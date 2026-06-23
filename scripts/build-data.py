import argparse
import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils.datetime import from_excel


REQUIRED_HEADERS = {
    "lot": "批次序号",
    "product": "货品名字",
    "bl": "提单号 (BL#)",
    "containers": "柜数",
    "pallets": "托数",
    "quantity": "进/出口件数",
    "net_weight": "发沪净重 (KG)",
    "gross_weight": "发沪毛重 (KG)",
    "cost": "成本总计 (USD)",
    "sales": "卖价总计 (USD)",
    "europe_etd": "欧洲ETD",
    "transshipment": "中转港及 ETA",
    "shanghai_eta": "上海ETA",
    "remark": "最新物流与节点状态",
}


def clean_text(value):
    if value is None:
        return ""
    text = re.sub(r"<br\s*/?>", " / ", str(value), flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


def normalize_header(value):
    return re.sub(r"\s+", "", clean_text(value)).lower()


def to_number(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?[\d,]+(?:\.\d+)?", clean_text(value).replace("$", ""))
    return float(match.group(0).replace(",", "")) if match else 0.0


def compact_number(value):
    number = to_number(value)
    return int(number) if number.is_integer() else number


def date_value(value, epoch):
    if value in (None, "", "待定", "-"):
        return "待定"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)) and value > 30000:
        return from_excel(value, epoch).strftime("%Y-%m-%d")
    return clean_text(value)


def parse_date(value):
    if not value or value == "待定":
        return None
    match = re.search(r"(20\d{2})-(\d{2})-(\d{2})", value)
    if match:
        return date(*map(int, match.groups()))
    match = re.search(r"(\d{1,2})月(\d{1,2})日", value)
    if match:
        month, day = map(int, match.groups())
        return date(date.today().year, month, day)
    return None


def transshipment_parts(value):
    text = clean_text(value) or "待定"
    if text == "待定":
        return "待定", "待定"
    port = text.split("/")[0].strip()
    match = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    eta = f"{date.today().year}-{int(match.group(1)):02d}-{int(match.group(2)):02d}" if match else "待定"
    return port, eta


def build_column_map(header_row):
    available = {normalize_header(value): index for index, value in enumerate(header_row) if value}
    result = {}
    missing = []
    for field, expected in REQUIRED_HEADERS.items():
        key = normalize_header(expected)
        if key not in available:
            missing.append(expected)
        else:
            result[field] = available[key]
    if missing:
        raise ValueError("Excel 缺少必要表头: " + ", ".join(missing))
    return result


def stage_for(remark, transshipment_eta, shanghai_eta):
    today = date.today()
    if "已抵上海" in remark or "已经抵达上海" in remark:
        return "arrived_shanghai", "抵沪"
    shanghai_date = parse_date(shanghai_eta)
    if shanghai_date and shanghai_date <= today:
        return "shanghai_due", "到港待确认"
    trans_date = parse_date(transshipment_eta)
    if trans_date and trans_date <= today:
        return "transshipment", "中转/二程"
    if "泰国" in remark or "改港" in remark:
        return "transshipment", "中转处理中"
    return "first_leg", "头程/待排"


def build_payload(input_path):
    workbook = openpyxl.load_workbook(input_path, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if len(rows) < 3:
        raise ValueError("Excel 没有明细数据")

    columns = build_column_map(rows[1])
    records = []
    shared_containers = {}
    normal_container_total = 0

    for row_number, row in enumerate(rows[2:], start=3):
        lot_raw = clean_text(row[columns["lot"]])
        if not lot_raw or lot_raw == "全局合计":
            continue

        product = clean_text(row[columns["product"]]) or "未注明货品"
        lot_match = re.search(r"Lot\s*([0-9]+(?:\.[0-9]+)?)", lot_raw, re.IGNORECASE)
        lot_id = lot_match.group(1) if lot_match else lot_raw
        container_text = clean_text(row[columns["containers"]])
        container_count = compact_number(container_text)
        is_shared = "共享" in container_text
        if is_shared:
            shared_containers[lot_id] = max(shared_containers.get(lot_id, 0), container_count)
        else:
            normal_container_total += container_count

        transshipment = clean_text(row[columns["transshipment"]])
        transshipment_port, transshipment_eta = transshipment_parts(transshipment)
        europe_etd = date_value(row[columns["europe_etd"]], workbook.epoch)
        shanghai_eta = date_value(row[columns["shanghai_eta"]], workbook.epoch)
        remark = clean_text(row[columns["remark"]]) or "待更新"
        stage, status = stage_for(remark, transshipment_eta, shanghai_eta)
        cost = round(to_number(row[columns["cost"]]), 2)
        sales = round(to_number(row[columns["sales"]]), 2)

        records.append({
            "id": f"{lot_id}-{product}",
            "lot": lot_id,
            "lotNumber": float(lot_id),
            "product": product,
            "bl": clean_text(row[columns["bl"]]) or "待定",
            "containersDisplay": container_text or "待定",
            "containerCount": container_count,
            "sharedContainers": is_shared,
            "pallets": compact_number(row[columns["pallets"]]),
            "quantity": clean_text(row[columns["quantity"]]) or "-",
            "netWeightKg": compact_number(row[columns["net_weight"]]),
            "grossWeightKg": compact_number(row[columns["gross_weight"]]),
            "costAmount": cost,
            "salesAmount": sales,
            "grossProfit": round(sales - cost, 2),
            "grossMargin": round((sales - cost) / sales * 100, 1) if sales else 0,
            "europeEtd": europe_etd,
            "transshipmentPort": transshipment_port,
            "transshipmentEta": transshipment_eta,
            "transshipmentDisplay": transshipment,
            "shanghaiEta": shanghai_eta,
            "stage": stage,
            "status": status,
            "remark": remark,
            "sourceRow": row_number,
        })

    records.sort(key=lambda item: (item["lotNumber"], item["product"]))
    lot_groups = {}
    for record in records:
        group = lot_groups.setdefault(record["lot"], {
            "lot": record["lot"], "costAmount": 0, "salesAmount": 0, "grossProfit": 0
        })
        group["costAmount"] += record["costAmount"]
        group["salesAmount"] += record["salesAmount"]
        group["grossProfit"] += record["grossProfit"]

    chart_groups = []
    for lot_id in sorted(lot_groups, key=float):
        group = lot_groups[lot_id]
        group["costAmount"] = round(group["costAmount"], 2)
        group["salesAmount"] = round(group["salesAmount"], 2)
        group["grossProfit"] = round(group["grossProfit"], 2)
        chart_groups.append(group)

    total_cost = round(sum(item["costAmount"] for item in records), 2)
    total_sales = round(sum(item["salesAmount"] for item in records), 2)
    total_profit = round(total_sales - total_cost, 2)
    total_containers = int(normal_container_total + sum(shared_containers.values()))

    return {
        "project": "Project B",
        "route": "Hungary → Transshipment → Shanghai",
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "sourceFile": Path(input_path).name,
        "summary": {
            "recordCount": len(records),
            "lotCount": len({item["lot"] for item in records}),
            "containers": total_containers,
            "pallets": int(sum(to_number(item["pallets"]) for item in records)),
            "costAmount": total_cost,
            "salesAmount": total_sales,
            "grossProfit": total_profit,
            "grossMargin": round(total_profit / total_sales * 100, 1) if total_sales else 0,
        },
        "chartGroups": chart_groups,
        "records": records,
    }


def main():
    parser = argparse.ArgumentParser(description="Build Project B dashboard data from Excel")
    parser.add_argument("--input", default="data/project-b-loading-details.xlsx")
    parser.add_argument("--output", default="data/lots.json")
    args = parser.parse_args()
    payload = build_payload(args.input)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output} with {len(payload['records'])} records")


if __name__ == "__main__":
    main()
