# Migration Map

> 规则：仅复制到新工程；旧目录不删除，保证可回滚。

| 旧路径 | 新路径 | 说明 |
|---|---|---|
| `backend_logic.py` | `src/backend_logic.py` | 核心逻辑 |
| `main.py` | `src/main.py` | 入口脚本 |
| `main_gui.py` | `src/main_gui.py` | GUI 入口 |
| `app_paths.py` | `src/app_paths.py` | 路径工具 |
| `archive.py` | `src/archive.py` | 归档逻辑 |
| `license_manager.py` | `src/license_manager.py` | 授权管理 |
| `verify_equivalence.py` | `src/verify_equivalence.py` | 校验脚本 |
| `kpi_tool/**` | `src/kpi_tool/**` | 子模块源码 |
| `KPI_Tool.spec` | `scripts/packaging/KPI_Tool.spec` | 主打包配置 |
| `活动保障指标自动化输出工具.spec` | `backup/legacy/活动保障指标自动化输出工具.spec` | 旧版打包配置备份 |
| `AUDIT.md` | `docs/audit/AUDIT.md` | 审计记录归档 |
| `logo.ico` | `assets/logo.ico` | 静态资源 |
| `.github/workflows/release-exe.yml` | `.github/workflows/release-exe.yml` | 发布流水线 |
| `网管指标/**` | `data_sample/**`（脱敏） | 原始生产样本不直接迁移 |
| `配置文件/项目配置.xlsx` | `configs/project_config.template.md` | 用示例模板替代 |
| `保障小区清单/保障小区清单.xlsx` | `assets/templates/保障小区清单.template.csv` | 用脱敏模板替代 |
| `build/dist/logs/历史数据` | 不迁移 | 运行产物默认排除 |

