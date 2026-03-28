from __future__ import annotations

import json
import sqlite3
import sys
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import auth
from catalog import Category, Item, load_catalog_from_path
from catalog_config import resolve_catalog_path
import records_db
from db_config import resolve_database_path
from login_ui import LoginScreen
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_NUM = ("Consolas", 11)


@dataclass(frozen=True)
class ValidLine:
    product_name: str
    quantity_str: str
    unit_price: Decimal
    total: Decimal

    def unit_price_db(self) -> str:
        return format(self.unit_price.quantize(Decimal("0.01")), "f")

    def total_price_db(self) -> str:
        return format(self.total.quantize(Decimal("0.01")), "f")


def _project_root() -> Path:
    return Path(__file__).resolve().parent


def _db_path() -> Path:
    return resolve_database_path(_project_root())


def fmt_money(d: Decimal) -> str:
    q = d.quantize(Decimal("0.01"))
    return f"¥ {q:.2f}"


_WIN_FILENAME_FORBIDDEN = frozenset('\\/:*?"<>|')


def _safe_export_filename_stem(username: str) -> str:
    """用户名 + 年月日时间，去掉 Windows 非法文件名字符（路径中不能用 /）。"""
    u = username.strip()
    safe_user = "".join(
        c if c not in _WIN_FILENAME_FORBIDDEN and ord(c) >= 32 else "_"
        for c in u
    )
    safe_user = safe_user.strip(". ") or "用户"
    now = datetime.now()
    # 对应「年-月-日」与「时间」，与导出时刻一致
    date_part = now.strftime("%Y-%m-%d")
    time_part = now.strftime("%H%M%S")
    return f"{safe_user}_{date_part}_{time_part}"


class ReconcileApp:
    def __init__(
        self,
        root: tk.Tk,
        categories: list[Category],
        db_path: Path,
        session: auth.UserSession,
        *,
        on_logout: Callable[[], None],
    ) -> None:
        self.root = root
        self.categories = categories
        self._category_by_name = {c.name: c for c in categories}
        self._db_path = db_path
        self._session = session
        self._logout_callback = on_logout
        records_db.ensure_schema(self._db_path)

        role_label = "管理员" if session.is_admin else "普通用户"
        root.title(f"对账小工具 — {session.username}（{role_label}）")
        root.minsize(380, 280)
        root.option_add("*Font", FONT_UI)

        main = ttk.Frame(root, padding=16)
        main.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        ttk.Label(main, text="商品分类").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.combo_cat = ttk.Combobox(
            main,
            state="readonly",
            values=[c.name for c in categories],
            width=32,
        )
        self.combo_cat.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.combo_cat.bind("<<ComboboxSelected>>", self._on_category)

        ttk.Label(main, text="商品").grid(row=2, column=0, sticky="w", pady=(0, 4))
        self.combo_item = ttk.Combobox(main, state="readonly", width=32)
        self.combo_item.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.combo_item.bind("<<ComboboxSelected>>", self._on_item)

        ttk.Label(main, text="数量").grid(row=4, column=0, sticky="w", pady=(0, 4))
        qty_frame = ttk.Frame(main)
        qty_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.entry_qty = ttk.Entry(qty_frame, width=16)
        self.entry_qty.pack(side="left")
        self.entry_qty.bind("<KeyRelease>", self._on_qty_change)
        self.entry_qty.bind("<FocusOut>", self._on_qty_change)

        box = ttk.LabelFrame(main, text="金额", padding=10)
        box.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="单价（可改）").grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.entry_unit = ttk.Entry(box, width=14, font=FONT_NUM, justify="right")
        self.entry_unit.grid(row=0, column=1, sticky="e")
        self.entry_unit.bind("<KeyRelease>", self._on_unit_change)
        self.entry_unit.bind("<FocusOut>", self._on_unit_change)

        ttk.Label(box, text="总价").grid(row=1, column=0, sticky="w", pady=(8, 0), padx=(0, 12))
        self.lbl_total = ttk.Label(box, text="—", anchor="e", font=FONT_NUM)
        self.lbl_total.grid(row=1, column=1, sticky="e", pady=(8, 0))

        btn_row = ttk.Frame(main)
        btn_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        btn_row.columnconfigure(0, weight=1)
        btn_left = ttk.Frame(btn_row)
        btn_left.grid(row=0, column=0, sticky="w")
        ttk.Button(btn_left, text="提交到本地数据库", command=self._on_submit).pack(
            side="left", padx=(0, 8)
        )
        self._btn_export = ttk.Button(
            btn_left, text="导出 Excel", command=self._on_export_excel
        )
        self._btn_export.pack(side="left", padx=(0, 8))
        ttk.Button(
            btn_left, text="清空历史数据", command=self._on_clear_history
        ).pack(side="left", padx=(0, 8))
        if session.is_admin:
            ttk.Button(
                btn_left, text="批量导入 Excel", command=self._on_import_excel
            ).pack(side="left")
        ttk.Button(btn_row, text="退出登录", command=self._handle_logout).grid(
            row=0, column=1, sticky="e", padx=(12, 0)
        )

        self.status = ttk.Label(main, text="", foreground="#666666")
        self.status.grid(row=8, column=0, columnspan=2, sticky="w")

        main.columnconfigure(0, weight=1)

        if categories:
            self.combo_cat.current(0)
            self._on_category()

    def _on_category(self, _event: object | None = None) -> None:
        name = self.combo_cat.get()
        cat = self._category_by_name.get(name)
        self.combo_item.set("")
        self.entry_unit.delete(0, tk.END)
        self.lbl_total.config(text="—")
        self.status.config(text="")
        self.entry_qty.delete(0, tk.END)

        if not cat or not cat.items:
            self.combo_item.config(values=[])
            return

        self.combo_item.config(values=[i.name for i in cat.items])
        self.combo_item.current(0)
        self._on_item()

    def _on_item(self, _event: object | None = None) -> None:
        item = self._selected_item()
        self.entry_unit.delete(0, tk.END)
        if item:
            self.entry_unit.insert(0, self._default_price_str(item.unit_price))
        self._update_totals()

    def _on_qty_change(self, _event: object | None = None) -> None:
        self._update_totals()

    def _on_unit_change(self, _event: object | None = None) -> None:
        self._update_totals()

    @staticmethod
    def _default_price_str(price: Decimal) -> str:
        return format(price.quantize(Decimal("0.01")), "f")

    def _selected_item(self) -> Item | None:
        cname = self.combo_cat.get()
        iname = self.combo_item.get()
        cat = self._category_by_name.get(cname)
        if not cat or not iname:
            return None
        for it in cat.items:
            if it.name == iname:
                return it
        return None

    def _compute_line(self) -> ValidLine | None:
        item = self._selected_item()
        if not item:
            self.entry_unit.delete(0, tk.END)
            self.lbl_total.config(text="—")
            self.status.config(text="")
            return None

        raw_u = self.entry_unit.get().strip()
        if not raw_u:
            self.status.config(text="请输入单价。")
            self.lbl_total.config(text="—")
            return None

        try:
            unit_price = Decimal(raw_u)
        except (InvalidOperation, ValueError):
            self.status.config(text="请输入有效单价（数字，可含小数）。")
            self.lbl_total.config(text="—")
            return None

        if unit_price < 0:
            self.status.config(text="单价不能为负数。")
            self.lbl_total.config(text="—")
            return None

        raw_q = self.entry_qty.get().strip()
        if not raw_q:
            self.status.config(text="")
            self.lbl_total.config(text="—")
            return None

        try:
            qty = Decimal(raw_q)
        except (InvalidOperation, ValueError):
            self.status.config(text="请输入有效数量（非负数字，可含小数）。")
            self.lbl_total.config(text="—")
            return None

        if qty < 0:
            self.status.config(text="数量不能为负数。")
            self.lbl_total.config(text="—")
            return None

        self.status.config(text="")
        total = (unit_price * qty).quantize(Decimal("0.01"))
        self.lbl_total.config(text=fmt_money(total))
        iname = self.combo_item.get().strip()
        return ValidLine(
            product_name=iname,
            quantity_str=raw_q,
            unit_price=unit_price,
            total=total,
        )

    def _update_totals(self) -> None:
        self._compute_line()

    def _on_submit(self) -> None:
        line = self._compute_line()
        if line is None:
            messagebox.showwarning(
                "无法提交",
                "请先选择商品，并填写有效的单价与数量。",
            )
            return

        ts = datetime.now().replace(microsecond=0).isoformat(sep=" ")
        try:
            records_db.insert_line_item(
                self._db_path,
                product_name=line.product_name,
                quantity=line.quantity_str,
                unit_price=line.unit_price_db(),
                total_price=line.total_price_db(),
                submitted_at=ts,
                submitted_by=self._session.username,
            )
        except OSError as e:
            messagebox.showerror("保存失败", str(e))
        except sqlite3.Error as e:
            messagebox.showerror("数据库错误", str(e))
        else:
            messagebox.showinfo("已保存", "记录已写入本地数据库。")

    def _on_export_excel(self) -> None:
        try:
            from excel_export import export_database_to_xlsx
        except ImportError as e:
            messagebox.showerror(
                "缺少依赖",
                "导出 Excel 需要先安装 openpyxl：\n"
                "pip install openpyxl\n\n"
                f"详情：{e}",
            )
            return

        default = f"{_safe_export_filename_stem(self._session.username)}.xlsx"
        path_str = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出为 Excel",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 工作簿", "*.xlsx"),
                ("所有文件", "*.*"),
            ],
            initialfile=default,
        )
        if not path_str:
            return

        out = Path(path_str)
        try:
            n = export_database_to_xlsx(
                self._db_path,
                out,
                viewer_username=self._session.username,
                is_admin=self._session.is_admin,
            )
        except (OSError, sqlite3.Error) as e:
            messagebox.showerror("导出失败", str(e))
        else:
            scope = "全部用户提交" if self._session.is_admin else "仅您提交的"
            messagebox.showinfo(
                "导出完成",
                f"已导出 {n} 条记录（{scope}）。\n\n{out}",
            )

    def _on_clear_history(self) -> None:
        if self._session.is_admin:
            if not messagebox.askyesno(
                "确认清空历史数据",
                "将删除本地数据库中「所有用户」的对账明细（含导入合并的数据），"
                "此操作不可恢复。\n\n确定要继续吗？",
            ):
                return
            try:
                n = records_db.delete_all_line_items(self._db_path)
            except (OSError, sqlite3.Error) as e:
                messagebox.showerror("清空失败", str(e))
                return
        else:
            if not messagebox.askyesno(
                "确认清空历史数据",
                f"将删除您（{self._session.username}）在本地提交的全部对账记录，"
                "此操作不可恢复。\n\n确定要继续吗？",
            ):
                return
            try:
                n = records_db.delete_line_items_for_submitter(
                    self._db_path, self._session.username
                )
            except (OSError, sqlite3.Error) as e:
                messagebox.showerror("清空失败", str(e))
                return
        messagebox.showinfo("已清空", f"已删除 {n} 条对账记录。")

    def _on_import_excel(self) -> None:
        if not self._session.is_admin:
            return
        try:
            from excel_export import merge_xlsx_into_database
        except ImportError as e:
            messagebox.showerror(
                "缺少依赖",
                "导入 Excel 需要先安装 openpyxl：\n"
                "pip install openpyxl\n\n"
                f"详情：{e}",
            )
            return

        paths_tuple = filedialog.askopenfilenames(
            parent=self.root,
            title="批量导入 Excel 合并到本地库（可多选）",
            filetypes=[
                ("Excel 工作簿", "*.xlsx"),
                ("所有文件", "*.*"),
            ],
        )
        paths = list(paths_tuple)
        if not paths:
            return

        total_inserted = 0
        row_errors: list[str] = []
        file_errors: list[str] = []
        ok_files = 0

        for path_str in paths:
            src = Path(path_str)
            name = src.name
            try:
                n, errs = merge_xlsx_into_database(self._db_path, src)
            except ValueError as e:
                file_errors.append(f"{name}：{e}")
                continue
            except (OSError, sqlite3.Error) as e:
                file_errors.append(f"{name}：{e}")
                continue
            ok_files += 1
            total_inserted += n
            for er in errs:
                row_errors.append(f"{name} — {er}")

        if ok_files == 0:
            max_fe = 10
            tail_fe = (
                f"\n… 另有 {len(file_errors) - max_fe} 个文件未列出"
                if len(file_errors) > max_fe
                else ""
            )
            messagebox.showerror(
                "导入失败",
                "所选文件均未能导入：\n\n"
                + "\n".join(file_errors[:max_fe])
                + tail_fe,
            )
            return

        hint = (
            "提示：同一文件再次导入会追加重复明细，请按需备份数据库。"
        )
        fail_part = (
            f"\n{len(file_errors)} 个文件未能导入（见后续提示）。"
            if file_errors
            else ""
        )
        messagebox.showinfo(
            "导入完成",
            f"已处理 {ok_files} 个文件（共选择 {len(paths)} 个），"
            f"累计合并 {total_inserted} 条记录到本地数据库。"
            f"{fail_part}\n\n{hint}",
        )
        if file_errors:
            max_fe = 6
            tail_fe = (
                f"\n… 另有 {len(file_errors) - max_fe} 个文件未列出"
                if len(file_errors) > max_fe
                else ""
            )
            messagebox.showwarning(
                "部分文件未导入",
                "以下文件未写入数据库：\n\n"
                + "\n".join(file_errors[:max_fe])
                + tail_fe,
            )
        if row_errors:
            max_show = 8
            tail = (
                f"\n… 另有 {len(row_errors) - max_show} 条错误未显示"
                if len(row_errors) > max_show
                else ""
            )
            body = "\n".join(row_errors[:max_show]) + tail
            messagebox.showwarning(
                "部分行已跳过",
                f"以下行未写入（共 {len(row_errors)} 条）：\n\n{body}",
            )

    def _handle_logout(self) -> None:
        self._logout_callback()


def main() -> None:
    path = resolve_catalog_path(_project_root())
    try:
        categories = load_catalog_from_path(path)
    except ImportError as e:
        messagebox.showerror("缺少依赖", str(e))
        sys.exit(1)
    except FileNotFoundError:
        messagebox.showerror(
            "无法加载数据",
            f"未找到商品目录文件：\n{path}\n\n"
            "默认从 data/products_catalog.xlsx 读取商品目录；也可在 data/app_config.json 中设置 "
            "catalog_path 指向其他文件，并保证文件存在。",
        )
        sys.exit(1)
    except json.JSONDecodeError as e:
        messagebox.showerror("JSON 解析失败", str(e))
        sys.exit(1)
    except ValueError as e:
        messagebox.showerror("数据格式错误", str(e))
        sys.exit(1)
    except OSError as e:
        messagebox.showerror("读取失败", f"无法读取商品目录文件：\n{e}")
        sys.exit(1)

    db_path = _db_path()
    records_db.ensure_schema(db_path)
    auth.ensure_users_table(db_path)

    root = tk.Tk()
    root.minsize(360, 420)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    def open_main(session: auth.UserSession) -> None:
        for w in root.winfo_children():
            w.destroy()
        ReconcileApp(
            root,
            categories,
            db_path,
            session,
            on_logout=show_login,
        )

    def show_login() -> None:
        root.title("对账小工具 — 登录")
        for w in root.winfo_children():
            w.destroy()
        screen = LoginScreen(root, db_path, on_success=open_main)
        screen.grid(row=0, column=0, sticky="nsew")

    show_login()
    root.mainloop()


if __name__ == "__main__":
    main()
