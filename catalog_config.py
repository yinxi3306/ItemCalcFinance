"""解析商品目录文件路径（JSON 或 Excel）。

优先级（从高到低）：
1. 环境变量 ITEMCALC_CATALOG（绝对路径，或相对项目根的路径）
2. 项目内 data/app_config.json 的 catalog_path 字段（同上）
3. 默认：<项目根>/data/products_catalog.xlsx（商品与分类从此 Excel 读取）

工作表名默认为「商品目录」；首行表头：分类、商品名、单价。
仍可通过 catalog_path 或环境变量改为 .json 等其他文件。
"""

from __future__ import annotations

import json
import os
from pathlib import Path

ENV_CATALOG = "ITEMCALC_CATALOG"
CONFIG_REL = Path("data") / "app_config.json"
DEFAULT_CATALOG_REL = Path("data") / "products_catalog.xlsx"


def _expand_path(raw: str, project_root: Path) -> Path:
    expanded = Path(os.path.expandvars(raw.strip())).expanduser()
    if not expanded.is_absolute():
        expanded = project_root / expanded
    return expanded.resolve()


def resolve_catalog_path(project_root: Path) -> Path:
    env_val = os.environ.get(ENV_CATALOG, "").strip()
    if env_val:
        return _expand_path(env_val, project_root)
    cfg_path = project_root / CONFIG_REL
    if cfg_path.is_file():
        try:
            data = json.loads(cfg_path.read_text(encoding="utf-8"))
            raw = (data.get("catalog_path") or "").strip()
            if raw:
                return _expand_path(raw, project_root)
        except (json.JSONDecodeError, OSError, TypeError):
            pass
    return (project_root / DEFAULT_CATALOG_REL).resolve()
