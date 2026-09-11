# Windows 打包说明

1. 将本目录上传到 GitHub 仓库，确保 `makro_scraper.py`、`requirements.txt` 和 `.github/workflows/build-windows.yml` 一起提交。
2. 在 GitHub 仓库的 **Actions** 页面运行 **Build Windows application**，或向 `main` / `master` 分支推送代码自动触发。
3. 构建完成后，在该次运行页面的 **Artifacts** 下载 `MakroScraper-windows`。
4. 解压后运行 `MakroScraper.exe`。程序会把生成的 Excel 保存到 exe 所在目录，文件名为 `makro_年月日_时分秒.xlsx`。

构建使用 Python 3.12、Windows 2022 runner，生成的程序适用于 Windows 10 及以上版本。当前不需要 GitHub Secret 或 Makro 账号。
