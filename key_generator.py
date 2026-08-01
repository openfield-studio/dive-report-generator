# -*- coding: utf-8 -*-
"""
開発者側で使うライセンスキー発行ツール。
配布先PCから伝えられた「このPCのID」を入力すると、そのPC専用の
ライセンスキーを計算して表示する。

秘密鍵ファイル private_key.hex がこのスクリプトと同じフォルダに必要。
このファイルは配布先PCには絶対にコピーしないこと。

使い方:
    python key_generator.py
"""

import os
import tkinter as tk
from tkinter import messagebox

import license_core

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PRIVATE_KEY_PATH = os.path.join(SCRIPT_DIR, "private_key.hex")


def main():
    root = tk.Tk()
    root.title("ライセンスキー発行ツール")
    root.resizable(False, False)

    pad = {"padx": 12, "pady": 6}

    tk.Label(root, text="配布先PCの「このPCのID」を入力してください。").grid(
        row=0, column=0, columnspan=2, sticky="w", **pad
    )

    tk.Label(root, text="PC ID").grid(row=1, column=0, sticky="w", **pad)
    id_entry = tk.Entry(root, width=48)
    id_entry.grid(row=1, column=1, **pad)
    id_entry.focus_set()

    tk.Label(root, text="発行キー").grid(row=2, column=0, sticky="w", **pad)
    key_entry = tk.Entry(root, width=48)
    key_entry.config(state="readonly")
    key_entry.grid(row=2, column=1, **pad)

    def generate():
        machine_id = id_entry.get().strip()
        if not machine_id:
            return
        if not os.path.exists(PRIVATE_KEY_PATH):
            messagebox.showerror(
                "秘密鍵が見つかりません",
                f"{PRIVATE_KEY_PATH}\nが見つかりません。private_key.hex をこのフォルダに置いてください。",
            )
            return
        key = license_core.compute_license_key(machine_id, PRIVATE_KEY_PATH)
        key_entry.config(state="normal")
        key_entry.delete(0, tk.END)
        key_entry.insert(0, key)
        key_entry.config(state="readonly")

    def copy_key():
        key = key_entry.get()
        if key:
            root.clipboard_clear()
            root.clipboard_append(key)

    btn_frame = tk.Frame(root)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
    tk.Button(btn_frame, text="キー生成", width=12, command=generate).pack(side="left", padx=6)
    tk.Button(btn_frame, text="コピー", width=12, command=copy_key).pack(side="left", padx=6)

    root.bind("<Return>", lambda e: generate())
    root.mainloop()


if __name__ == "__main__":
    main()
