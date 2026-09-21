from __future__ import annotations

from pathlib import Path

from daily_sales_app.config_store import ConfigStore
from daily_sales_app.gui import DailySalesApp


def main() -> None:
    app_dir = Path(__file__).resolve().parent
    store = ConfigStore(app_dir / "app_data" / "state.json")
    app = DailySalesApp(store=store, app_dir=app_dir)
    app.mainloop()


if __name__ == "__main__":
    main()
