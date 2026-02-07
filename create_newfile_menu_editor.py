import sys
import tkinter as tk
from tkinter import BOTH, LEFT, RIGHT, Y, Frame, messagebox, simpledialog
import winreg
import ctypes
import threading

# 管理者チェック
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class ShellNewManager:

    IGNORE_VALUE = (".contact", ".library-ms", ".lnk", ".zip")

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("新規作成メニュー カスタマイザー")
        self.root.geometry("500x400")

        frame = Frame(self.root)
        frame.pack(pady=10)

        scrbar = tk.Scrollbar(frame)
        scrbar.pack(side=RIGHT, fill=Y)

        self.listbox = tk.Listbox(frame, width=60, yscrollcommand=scrbar.set)
        self.listbox.pack(side=LEFT, fill=BOTH)

        scrbar.config(command=self.listbox.yview)

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=5)

        tk.Button(self.btn_frame, text="有効/無効切替", command=self.toggle_item).grid(row=0, column=0, padx=5)
        tk.Button(self.btn_frame, text="追加", command=self.add_item).grid(row=0, column=1, padx=5)
        tk.Button(self.btn_frame, text="削除", command=self.delete_item).grid(row=0, column=2, padx=5)
        tk.Button(self.btn_frame, text="再読み込み", command=self.update_items).grid(row=0, column=3, padx=5)

        self.status = tk.Label(self.root, text="", fg="gray")
        self.status.pack()

        self.root.update()
        self.update_items()
        self.root.mainloop()

    def traverse_registry(self, root, path):
        try:
            with winreg.OpenKey(root, path) as key:
                index = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, index)
                        if subkey_name in self.IGNORE_VALUE:
                            index += 1
                            continue

                        subkey_path = f"{path + '\\' if path else ''}{subkey_name}"

                        for suffix, enabled in [("ShellNew", True), ("ShellNew-", False)]:
                            try:
                                with winreg.OpenKey(root, f"{subkey_path}\\{suffix}"):
                                    filetype = subkey_path.split("\\")[0]
                                    if filetype.startswith("."):
                                        display_name = self.get_display_name_from_type(filetype)
                                        self.items.append((subkey_path, filetype, display_name, enabled))
                            except FileNotFoundError:
                                pass

                        self.traverse_registry(root, subkey_path)
                        index += 1
                    except OSError:
                        break
        except FileNotFoundError:
            pass

    def get_display_name_from_type(self, filetype):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, filetype) as key:
                name, _ = winreg.QueryValueEx(key, "")
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name) as subkey:
                    display_name, _ = winreg.QueryValueEx(subkey, "")
                return display_name
        except:
            return "不明"

    def load_items(self):
        self.items_state = False
        self.listbox.delete(0, tk.END)
        self.items = []

        try:
            self.traverse_registry(winreg.HKEY_CLASSES_ROOT, "")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

        self.items.sort(key=lambda x: x[1])
        for i, (_, ext, name, enabled) in enumerate(self.items):
            state = "✔" if enabled else "✖"
            self.listbox.insert(tk.END, f"{state} {ext} ({name})")
            if enabled:
                self.listbox.itemconfig(i, {'fg': 'green'})
            else:
                self.listbox.itemconfig(i, {'fg': 'red'})

        self.items_state = True

    def update_items(self):
        self.animate_status()
        threading.Thread(target=self.load_items, daemon=True).start()

        def check():
            if self.items_state:
                self.stop_status()
                return
            self.root.after(100, check)

        self.root.after(20, check)

    def toggle_item(self):
        selection = self.listbox.curselection()
        if not selection:
            return
        index = selection[0]
        path, ext, name, enabled = self.items[index]

        base = f"{path}\\ShellNew"
        alt = f"{path}\\ShellNew-"

        try:
            if enabled:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, alt)  # 念のため先に削除
                winreg.RenameKey(winreg.HKEY_CLASSES_ROOT, base, "ShellNew-")
            else:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, base)  # 念のため先に削除
                winreg.RenameKey(winreg.HKEY_CLASSES_ROOT, alt, "ShellNew")

            self.status.config(text=f"{ext} を {'無効化' if enabled else '有効化'} しました")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

        self.update_items()

    def delete_item(self):
        def delete_key_recursive(root, path):
            try:
                with winreg.OpenKey(root, path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                    while True:
                        try:
                            subkey = winreg.EnumKey(key, 0)
                            delete_key_recursive(root, f"{path}\\{subkey}")
                        except OSError:
                            break
                winreg.DeleteKey(root, path)
            except FileNotFoundError:
                pass

        selection = self.listbox.curselection()
        if not selection:
            return
        index = selection[0]
        path, ext, name, _ = self.items[index]
        if messagebox.askyesno("確認", f"{name}({path}) を新規作成メニューから完全に削除しますか？"):
            try:
                for suffix in ["ShellNew", "ShellNew-"]:
                    delete_key_recursive(winreg.HKEY_CLASSES_ROOT, f"{path}\\{suffix}")
                self.status.config(text=f"{name}({path}) を削除しました")
                self.update_items()
            except Exception as e:
                messagebox.showerror("エラー", str(e))

    def add_item(self):
        ext = simpledialog.askstring("拡張子", "拡張子（例：.md）を入力してください：")
        if not ext or not ext.startswith("."):
            messagebox.showwarning("警告", "正しい拡張子を入力してください（例：.md）")
            return
        name = simpledialog.askstring("表示名", f"{ext} の表示名を入力してください：")
        try:
            ext_path = ext
            type_name = ext[1:] + "file"
            # 既に拡張子キーがあるかチェック
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ext_path):
                    ext_exists = True
            except FileNotFoundError:
                ext_exists = False

            if not ext_exists:
                with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, ext_path) as ext_key:
                    winreg.SetValueEx(ext_key, "", 0, winreg.REG_SZ, type_name)

                with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, type_name) as ft_key:
                    winreg.SetValueEx(ft_key, "", 0, winreg.REG_SZ, name)

            with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f"{ext_path}\\ShellNew") as sn_key:
                winreg.SetValueEx(sn_key, "NullFile", 0, winreg.REG_SZ, "")

            self.status.config(text=f"{name} を追加しました")
            self.update_items()
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def animate_status(self, state=0):
        text = ["更新中", ""]
        self.status.config(text=text[state], fg="blue")
        self.root.update()
        self._status_job = self.root.after(500, self.animate_status, not state)

    def stop_status(self, msg="完了"):
        if hasattr(self, '_status_job'):
            self.root.after_cancel(self._status_job)
        self.status.config(text=msg, fg="gray")
        self.root.update()

if __name__ == "__main__":
    if not is_admin():
        print("権限不足: 管理者として実行してください")
        sys.exit(-1)
    ShellNewManager()
