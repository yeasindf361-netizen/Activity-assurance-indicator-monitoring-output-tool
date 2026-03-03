# FAQ

## Q1: 同事下载后如何使用？
下载 Release 的 zip，解压后按 `快速使用说明.txt` 执行。

## Q2: 为什么仓库里没有 dist 和历史输出？
这些是运行产物，默认不入库，避免仓库膨胀和污染。

## Q3: 配置文件在哪里？
查看 `configs/` 下示例文件，复制后按实际环境填写。

## Q4: 打包失败怎么办？
1. 检查 Python 版本是否 3.13+  
2. 执行 `pip install -r requirements.txt`  
3. 检查 `scripts/packaging/KPI_Tool.spec` 路径是否存在。

