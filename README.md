# Project B 匈牙利至上海看板

Excel 数据源：`data/project-b-loading-details.xlsx`。

每次更新 Excel 后运行：

```powershell
python scripts/build-data.py
git add data/project-b-loading-details.xlsx data/lots.json
git commit -m "Update Project B details"
git push
```

GitHub Actions 会自动生成数据并发布 GitHub Pages。
