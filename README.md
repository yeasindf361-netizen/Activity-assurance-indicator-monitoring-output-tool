# Event Support Indicator Monitoring Tool

活动保障指标监控与自动化输出工具（标准工程版）。

## 1. 项目目标
- 支持 4G/5G 指标读取、计算、通报输出。
- 支持可维护开发（源码/配置/脚本/文档分层）。
- 支持 GitHub Release 分发给同事直接下载使用。

## 2. 目录结构
```text
src/            核心源码
configs/        配置模板（示例）
assets/         资源与模板
scripts/        本地运行/打包/清理脚本
docs/           说明文档、FAQ、维护手册
data_sample/    脱敏样例数据
dist/           打包产物（不入库）
```

## 3. 环境要求
- Windows 10/11
- Python 3.13+

## 4. 安装依赖
```powershell
pip install -r requirements.txt
```

## 5. 本地运行
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_local.ps1
```

## 6. 一键打包
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_release.ps1 -Version v1.0.0
```

产物输出：
- `dist/Event-Support-Tool-v1.0.0-win64.zip`

## 7. 输入输出说明
- 输入：
  - 配置模板：`configs/`
  - 样例模板：`assets/templates/`
- 输出：
  - 运行产物默认写入运行目录下输出路径（已通过 `.gitignore` 排除）。

## 8. GitHub Release 使用（给同事）
1. 打开仓库 `Releases` 页面。
2. 下载最新 `Event-Support-Tool-<tag>-win64.zip`。
3. 解压后按 `快速使用说明.txt` 操作。

## 9. 版本管理
- 使用语义化版本：`vMAJOR.MINOR.PATCH`
- 变更记录见 `CHANGELOG.md`

## 10. 常见问题
详见 `docs/FAQ.md`。

## 11. 升级方式
1. 拉取最新源码。
2. 查看 `CHANGELOG.md` 与 `docs/MAINTENANCE.md`。
3. 重新打包并发布新 Release。

