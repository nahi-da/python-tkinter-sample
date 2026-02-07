import tkinter as tk
from tkinter import ttk

root = tk.Tk()

table = tk.Frame(root)
table.pack(fill="both", expand=True, padx=10, pady=10)

headers = ["名前", "年齢", "種類", "メモ"]

# ヘッダ行
for col, text in enumerate(headers):
    lbl = tk.Label(table, text=text, anchor="center")
    lbl.grid(row=0, column=col, sticky="ew", padx=3, pady=3)

# 入力行
for row in range(1, 4):
    e1 = tk.Entry(table)
    e2 = tk.Entry(table)
    cb = ttk.Combobox(table, values=["A", "B", "C"])
    e3 = tk.Entry(table)

    e1.grid(row=row, column=0, sticky="ew", padx=3, pady=3)
    e2.grid(row=row, column=1, sticky="ew", padx=3, pady=3)
    cb.grid(row=row, column=2, sticky="ew", padx=3, pady=3)
    e3.grid(row=row, column=3, sticky="ew", padx=3, pady=3)

# 列幅制御（重要）
for i in range(4):
    table.grid_columnconfigure(i, weight=1)

root.mainloop()
