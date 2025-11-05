# Zemax Glasscat 2 TracePro

一键将 Zemax 的玻璃库（`*.AGF`）导入 TracePro 中

> [!IMPORTANT] 
> 仅在 `TracePro 7.x` 测试，不一定适用于其他版本，运行前请至少启动 TracePro 一次

## 使用方法

1. 安装 Python 3.8 或更高版本
2. 下载本仓库中的 `.py` 文件，右键使用 Python 打开（不是 IDLE）
3. 运行后会提示选择要导入的玻璃库，选择后会自动导入到 TracePro 中
4. 打开 TracePro 的材料库进行验证

## 注意事项

1. 如果导入不成功，可以通过在 CMD 里执行 `python zemax_glasscat_to_tracepro.py` 查看错误消息
2. 若导入的玻璃为 TracePro 自带的（例如 TracePro 中自带成都光明的玻璃库），TracePro 在直接打开 Zemax
   文件时会优选使用自带的而非导入的，需要手动修改，或者通过其他方法删除 TracePro 自带的成都光明的玻璃库
3. 脚本目前支持以下种公式，能覆盖大部分玻璃：
   - Schoot
   - Sellmeier 1
   - Herzberger
   - Sellmeier 2
   - Conrady
   - Sellmeier 3
   - Handbook 1
   - Handbook 2
   - Sellmeier 4
