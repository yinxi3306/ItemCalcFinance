from __future__ import annotations

import sqlite3
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, ttk

import auth
from auth import UserSession

FONT_UI = ("Microsoft YaHei UI", 10)


class LoginScreen(ttk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        db_path: Path,
        *,
        on_success: Callable[[UserSession], None],
    ) -> None:
        super().__init__(master, padding=24)
        self._db_path = db_path
        self._on_success = on_success

        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="对账小工具", font=("Microsoft YaHei UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 16)
        )
        ttk.Label(self, text="登录账号以继续使用（数据仅存本机）。").grid(
            row=1, column=0, columnspan=2, pady=(0, 16)
        )

        ttk.Label(self, text="用户名").grid(row=2, column=0, sticky="w", pady=(0, 4))
        self.entry_user = ttk.Entry(self, width=28)
        self.entry_user.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        ttk.Label(self, text="密码").grid(row=4, column=0, sticky="w", pady=(0, 4))
        self.entry_pwd = ttk.Entry(self, width=28, show="*")
        self.entry_pwd.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 16))

        btn_row = ttk.Frame(self)
        btn_row.grid(row=6, column=0, columnspan=2, sticky="ew")
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)

        ttk.Button(btn_row, text="登录", command=self._login).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(btn_row, text="注册", command=self._open_register).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        self.entry_user.bind("<Return>", lambda e: self._login())
        self.entry_pwd.bind("<Return>", lambda e: self._login())

        self.option_add("*Font", FONT_UI)

    def _login(self) -> None:
        user = self.entry_user.get()
        pwd = self.entry_pwd.get()
        try:
            session = auth.authenticate(self._db_path, user, pwd)
        except (OSError, sqlite3.Error) as e:
            messagebox.showerror("登录失败", str(e))
            return
        if session is None:
            messagebox.showwarning("登录失败", "用户名或密码不正确。")
            return
        self._finish(session)

    def _finish(self, session: UserSession) -> None:
        self.destroy()
        self._on_success(session)

    def _open_register(self) -> None:
        RegisterDialog(self.winfo_toplevel(), self._db_path, self._on_registered)

    def _on_registered(self, role: str) -> None:
        if role == auth.ROLE_ADMIN:
            messagebox.showinfo(
                "注册成功",
                "首位用户已自动设为管理员。请使用刚注册的账号登录。",
            )
        else:
            messagebox.showinfo("注册成功", "请使用新账号登录。")


class RegisterDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        db_path: Path,
        on_done: Callable[[str], None],
    ) -> None:
        super().__init__(master)
        self._db_path = db_path
        self._on_done = on_done
        self.title("注册账号")
        self.resizable(False, False)
        self.option_add("*Font", FONT_UI)

        f = ttk.Frame(self, padding=20)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="用户名（≥2 字符）").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.u = ttk.Entry(f, width=26)
        self.u.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(f, text="密码（≥6 位）").grid(row=2, column=0, sticky="w", pady=(0, 4))
        self.p1 = ttk.Entry(f, width=26, show="*")
        self.p1.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(f, text="确认密码").grid(row=4, column=0, sticky="w", pady=(0, 4))
        self.p2 = ttk.Entry(f, width=26, show="*")
        self.p2.grid(row=5, column=0, sticky="ew", pady=(0, 16))

        ttk.Button(f, text="创建账号", command=self._submit).grid(row=6, column=0, sticky="ew")

        self.transient(master)
        self.grab_set()
        self.u.focus_set()

    def _submit(self) -> None:
        if self.p1.get() != self.p2.get():
            messagebox.showwarning("注册失败", "两次输入的密码不一致。")
            return
        try:
            role = auth.register_user(self._db_path, self.u.get(), self.p1.get())
        except ValueError as e:
            messagebox.showwarning("注册失败", str(e))
            return
        except (OSError, sqlite3.Error) as e:
            messagebox.showerror("注册失败", str(e))
            return
        self.grab_release()
        self.destroy()
        self._on_done(role)
