from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


DEFAULT_STATE: dict[str, Any] = {
    "sku_mapping_path": "",
    "last_output_dir": "",
    "last_eur_to_cny_rate": 8.0,
    "platforms": [
        {"name": "CD", "operator": "陆美婷；陈嘉仪；卢善程", "requires_cost": True},
        {"name": "PM", "operator": "陆美婷；陈嘉仪；卢善程", "requires_cost": True},
        {"name": "Fnac", "operator": "陆美婷；陈嘉仪；卢善程", "requires_cost": True},
        {"name": "BOL", "operator": "杨可依", "requires_cost": False},
        {"name": "PCC", "operator": "杨可依", "requires_cost": False},
        {"name": "速卖通", "operator": "郑逸君", "requires_cost": False},
        {"name": "Miravia", "operator": "许悦娜", "requires_cost": False},
        {"name": "Otto", "operator": "许悦娜", "requires_cost": False},
    ],
}


class ConfigStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.state = self._load()

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return deepcopy(DEFAULT_STATE)
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return deepcopy(DEFAULT_STATE)

        state = deepcopy(DEFAULT_STATE)
        if isinstance(data, dict):
            state.update({k: v for k, v in data.items() if k in state})
        if not isinstance(state.get("platforms"), list) or not state["platforms"]:
            state["platforms"] = deepcopy(DEFAULT_STATE["platforms"])
        else:
            self._migrate_platforms(state["platforms"])
        return state

    @staticmethod
    def _migrate_platforms(platforms: list[dict[str, Any]]) -> None:
        shared_operator = "陆美婷；陈嘉仪；卢善程"
        for item in platforms:
            name = str(item.get("name", "")).strip()
            operator = str(item.get("operator", "")).strip()
            if name in {"CD", "PM", "Fnac"} and operator in {"陆美婷", "陈嘉仪", "卢善程"}:
                item["operator"] = shared_operator

    def save(self) -> None:
        self.path.write_text(
            json.dumps(self.state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def get_platforms(self) -> list[dict[str, Any]]:
        return list(self.state.get("platforms", []))
