# BOL_DailyOrder

Windows 桌面工具项目，用于按平台顺序上传每日订单文件、读取 SKU 对应表和成本报表，并生成销售汇总 Excel。

## 当前功能

- 按平台顺序逐个上传文件
- 支持 `上一步` 回退并重传
- `运营对应店铺` 页面可编辑平台、运营和是否需要成本报表
- SKU 对应表路径持久化保存
- 成本报表作为每日流程最后一步上传
- 未匹配 SKU 仅在界面展示 `user_account` 和 `refrence_no`
- 导出的 Excel 表头日期使用“程序运行日的前一天”

## 本地运行

```bash
pip install -r requirements.txt
python main.py
```

## GitHub Build Windows EXE

仓库已经带了 GitHub Actions 工作流：

- `.github/workflows/build-windows-exe.yml`

上传到 GitHub 后，你可以在仓库的 `Actions` 里手动触发 `Build Windows EXE`。

构建完成后，会生成一个 `windows-exe` artifact，里面包含：

- `BOL_DailyOrder.exe`

## 项目结构

```text
BOL_DailyOrder/
  .github/workflows/build-windows-exe.yml
  daily_sales_app/
    __init__.py
    config_store.py
    gui.py
    processor.py
  main.py
  requirements.txt
  README.md
  .gitignore
```
