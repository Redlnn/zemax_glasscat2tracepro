# Zemax Glasscat 2 TracePro

一键将 Zemax 的玻璃库（`*.AGF`）导入 TracePro 中

## 使用方法

1. 安装 Python 3.8 或更高版本
2. 下载本仓库中的 `.py` 文件，右键使用 Python 打开（不是 IDLE）
3. 运行后会提示选择要导入的玻璃库，选择后会自动导入到 TracePro 中
4. 打开 TracePro 的材料库进行验证

## 注意事项

1. 如果导入不成功，可以通过在 CMD 里执行 `python zemax_glasscat_to_tracepro.py` 查看错误消息
2. 若导入的玻璃为 TracePro 自带的（例如 TracePro 中自带成都光明的玻璃库），TracePro 在直接打开 Zemax
   文件时会优选使用自带的而非导入的，需要手动修改，或者通过其他方法删除 TracePro 自带的成都光明的玻璃库
3. 脚本目前仅支持以下 3 种公式，能覆盖大部分玻璃，若遇到无法导入或导入后缺少玻璃，请通过 **issue** 联系我
   - Schoot
   - Sellmeier 1
   - Herzberger
