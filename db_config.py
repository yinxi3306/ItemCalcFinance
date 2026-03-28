"""解析共用 SQLite 数据库路径。

优先级（从高到低）：
1. 环境变量 ITEMCALC_DB（绝对路径，或相对项目根的路径）
2. 项目内 data/app_config.json 的 database_path 字段（同上）
3. 默认：<项目根>/data/reconcile.db

部署示例（多 Windows 用户或固定位置）：
- 设置 ITEMCALC_DB=%ProgramData%\\ItemCalcFinance\\reconcile.db
- 或在 app_config.json 中写 \"database_path\": \"C:/.../reconcile.db\"

注意：多机同时写入网络盘上的同一 SQLite 文件存在锁定与损坏风险，请控制并发并做好备份。
"""

from __future__ import annotations

import json
import os
from pathlib import Path

ENV_DATABASE = "ITEMCALC_DB"
CONFIG_REL = Path("data") / "app_config.json"
DEFAULT_DB_REL = Path("data") / "reconcile.db"


def _expand_path(raw: str, project_root: Path) -> Path:
    expanded = Path(os.path.expandvars(raw.strip())).expanduser()
    if not expanded.is_absolute():
        expanded = project_root / expanded
    return expanded.resolve()


def resolve_database_path(project_root: Path) -> Path:
    env_val = os.environ.get(ENV_DATABASE, "").strip()
    if env_val:
        path = _expand_path(env_val, project_root)
    else:
        cfg_path = project_root / CONFIG_REL
        path = project_root / DEFAULT_DB_REL
        if cfg_path.is_file():
            try:
                data = json.loads(cfg_path.read_text(encoding="utf-8"))
                raw = (data.get("database_path") or "").strip()
                if raw:
                    path = _expand_path(raw, project_root)
            except (json.JSONDecodeError, OSError, TypeError):
                pass

    path.parent.mkdir(parents=True, exist_ok=True)
    return path
