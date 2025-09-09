
import csv
import os
import sys
import datetime
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

COLUMNS = [
    "รอบเดือน",
    "หมายเลขรายการ",
    "ชื่อSKU",
    "ชื่อสินค้า",
    "จำนวน (ชิ้น)",
    "ราคาสินค้าต่อหน่วย (บาท)",
    "ค่าส่งรวม (บาท)",
    "ต้นทุนรวม (บาท)",
    "ต้นทุนต่อตัว (บาท)",
    "Barcode",
    "Scan",
]

AUTOSAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_session.csv")
SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_settings.json")

class EditableTreeview(ttk.Treeview):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self._edit_widget = None
        self.last_anchor = ("#1", None)
        self.bind("<Double-1>", self._begin_edit_cell)
        self.bind("<Button-1>", self._on_single_click)

    def _on_single_click(self, event):
        region = self.identify("region", event.x, event.y)
        row_id = self.identify_row(event.y)
        col_id = self.identify_column(event.x)
        if region == "cell" and row_id and col_id:
            self.last_anchor = (col_id, row_id)
        if self._edit_widget is not None:
            self._end_edit_cell(save=True)

    def _begin_edit_cell(self, event):
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.identify_row(event.y)
        col_id = self.identify_column(event.x)
        if not row_id or not col_id:
            return
        self.last_anchor = (col_id, row_id)
        col_index = int(col_id[1:]) - 1
        if COLUMNS[col_index] == "Scan":
            return
        bbox = self.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.set(row_id, col_id)
        self._edit_row = row_id
        self._edit_col = col_id
        self._edit_widget = tk.Entry(self, borderwidth=1, bg="white", fg="black")
        self._edit_widget.insert(0, value)
        self._edit_widget.select_range(0, tk.END)
        self._edit_widget.focus_set()
        self._edit_widget.place(x=x, y=y, width=w, height=h)
        self._edit_widget.bind("<Return>", lambda e: self._end_edit_cell(True))
        self._edit_widget.bind("<Escape>", lambda e: self._end_edit_cell(False))
        self._edit_widget.bind("<FocusOut>", lambda e: self._end_edit_cell(True))

    def _end_edit_cell(self, save):
        if self._edit_widget is None:
            return
        new_value = self._edit_widget.get()
        if save:
            self.set(self._edit_row, self._edit_col, new_value)
            root = self._root()
            if hasattr(root, "autosave"):
                try:
                    root.autosave()
                except Exception:
                    pass
        self._edit_widget.destroy()
        self._edit_widget = None
        self._edit_row = None
        self._edit_col = None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("บัญชีต้นทำนำเข้าสินค้า + Stock in - out (Dawson DIY)")
        self.geometry("1280x720")
        self.minsize(1000, 560)

        # Apply Windows 98-ish colors
        self.configure(bg="#c0c0c0")
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TButton", background="#d4d0c8", foreground="black")
        style.map("TButton", background=[("active", "#ffffff")])
        style.configure("Treeview", background="white", fieldbackground="white", foreground="black")
        style.configure("TLabel", background="#c0c0c0", foreground="black")
        style.configure("TCheckbutton", background="#c0c0c0", foreground="black")
        style.configure("TEntry", fieldbackground="white", foreground="black")

        # Auto-Scan toggle
        self.auto_scan = tk.BooleanVar(value=True)

        self.auto_scan = tk.BooleanVar(value=True)

        self._load_settings()
        self._scan_debounce_id = None
        self._build_toolbar()
        self._build_table()
        self._build_scan_panel()
        self._build_statusbar()

        self._load_autosave_or_init()

        # Shortcuts
        self.bind_all("<Control-s>", lambda e: self.save_csv())
        self.bind_all("<Control-o>", lambda e: self.load_csv())
        self.bind_all("<Control-n>", lambda e: self.add_row_and_save())
        self.bind_all("<Delete>", lambda e: self.delete_selected_and_save())
        self.bind_all("<Control-v>", lambda e: self.paste_from_clipboard())
        self.bind_all("<Command-v>", lambda e: self.paste_from_clipboard())

        # ---------- Settings (persist last export folder) ----------
    def _load_settings(self):
        self.settings = {}
        try:
            if os.path.exists(SETTINGS_PATH):
                import json
                with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f) or {}
        except Exception:
            self.settings = {}

    def _save_settings(self):
        try:
            import json
            with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------- Auto-Scan Helpers ----------
    def _on_scan_var_write(self):
        # Debounce rapid key events from barcode scanner
        if not getattr(self, 'auto_scan', tk.BooleanVar(value=True)).get():
            return
        # cancel pending timer
        if getattr(self, '_scan_debounce_id', None):
            try:
                self.after_cancel(self._scan_debounce_id)
            except Exception:
                pass
        # schedule a short delay; if input stays stable, fire scan
        self._scan_debounce_id = self.after(150, self._auto_scan_fire)

    def _auto_scan_fire(self):
        try:
            data = self.scan_var.get().strip()
        except Exception:
            data = self.scan_entry.get().strip()
        if not data:
            return
        # simple heuristic: treat as scan if length >= 4
        if len(data) >= 4:
            self.process_scan_and_save()

# ---------- UI ----------
    def _build_toolbar(self):
        bar = tk.Frame(self, padx=6, pady=4, bg="#c0c0c0")
        bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(bar, text="เพิ่มแถว", command=self.add_row_and_save).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="ลบแถว", command=self.delete_selected_and_save).pack(side=tk.LEFT, padx=2)

        ttk.Button(bar, text="Load CSV", command=self.load_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="Save CSV", command=self.save_csv).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="Paste", command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=8)

        ttk.Button(bar, text="Gen Barcode", command=self.generate_barcode_for_selected_or_empty_and_save).pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="Recalc", command=self.recalc_all_and_save).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="Clear Scan", command=self.clear_scan_selected_and_save).pack(side=tk.LEFT, padx=8)

        ttk.Button(bar, text="Export Barcodes (PNG)", command=lambda: self.export_barcodes_png(selected_only=True)).pack(side=tk.LEFT, padx=8)

    def _build_table(self):
        container = tk.Frame(self, bg="#c0c0c0")
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tree = EditableTreeview(container, columns=COLUMNS, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150 if col == "ชื่อสินค้า" else 140, anchor="w")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

    def _build_scan_panel(self):
        pane = tk.Frame(self, padx=6, pady=6, bg="#c0c0c0")
        pane.pack(side=tk.TOP, fill=tk.X)
        ttk.Checkbutton(pane, text="Auto-Scan", variable=self.auto_scan, command=self._maybe_focus_scan).pack(side=tk.LEFT, padx=2)
        ttk.Label(pane, text="ช่องสแกน:").pack(side=tk.LEFT, padx=6)
        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(pane, width=40, textvariable=self.scan_var)
        self.scan_entry.pack(side=tk.LEFT, padx=2)
        self.scan_entry.bind("<Return>", lambda e: self.process_scan_and_save())
        # Auto-Scan: trigger when input stops changing briefly
        try:
            self.scan_var.trace_add("write", lambda *args: self._on_scan_var_write())
        except Exception:
            pass
        ttk.Button(pane, text="Scan", command=self.process_scan_and_save).pack(side=tk.LEFT, padx=6)

    def _build_statusbar(self):
        self.status = tk.StringVar(value="พร้อมใช้งาน")
        bar = tk.Label(self, textvariable=self.status, anchor="w", bg="#808080", fg="black", relief=tk.SUNKEN)
        bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _set_status(self, text: str):
        try:
            self.status.set(text)
        except Exception:
            pass

    # ---------- Autosave ----------
    def autosave(self):
        try:
            self._save_to_path(AUTOSAVE_PATH)
        except Exception:
            pass

    def _load_autosave_or_init(self):
        if os.path.exists(AUTOSAVE_PATH):
            try:
                self._load_from_path(AUTOSAVE_PATH)
                self._set_status("โหลดข้อมูลจาก last_session.csv แล้ว")
                return
            except Exception:
                pass
        self.add_row()
        self.autosave()

    # ---------- Data Ops ----------
    def add_row(self):
        values = [""] * len(COLUMNS)
        values[0] = datetime.datetime.now().strftime("%Y-%m")
        self.tree.insert("", "end", values=values)
        self._set_status("เพิ่มแถวใหม่แล้ว")

    def add_row_and_save(self):
        self.add_row()
        self.autosave()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            self._set_status("ยังไม่ได้เลือกแถว")
            return
        for iid in sel:
            self.tree.delete(iid)
        self._set_status(f"ลบ {len(sel)} แถวแล้ว")

    def delete_selected_and_save(self):
        self.delete_selected()
        self.autosave()

    def load_csv(self):
        path = filedialog.askopenfilename(title="เลือกไฟล์ CSV", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            self._load_from_path(path)
            self.autosave()
            self._set_status(f"โหลดข้อมูลจาก {os.path.basename(path)} แล้ว")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    def _load_from_path(self, path: str):
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            missing = [c for c in COLUMNS if c not in reader.fieldnames]
            if missing:
                raise ValueError("คอลัมน์หายไป: " + ", ".join(missing))
            for iid in self.tree.get_children():
                self.tree.delete(iid)
            for row in reader:
                self.tree.insert("", "end", values=[row.get(c, "") for c in COLUMNS])

    def save_csv(self):
        path = filedialog.asksaveasfilename(title="บันทึกเป็น CSV", defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            self._save_to_path(path)
            self._set_status(f"บันทึกข้อมูลไปที่ {os.path.basename(path)} แล้ว")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    def _save_to_path(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNS)
            writer.writeheader()
            for iid in self.tree.get_children():
                values = self.tree.item(iid, "values")
                row = {col: values[idx] for idx, col in enumerate(COLUMNS)}
                writer.writerow(row)

    # ---------- Paste ----------
    def paste_from_clipboard(self):
        try:
            raw = self.clipboard_get()
        except Exception:
            self._set_status("ไม่มีข้อมูลในคลิปบอร์ด")
            return
        rows = [r for r in raw.splitlines() if r.strip() != ""]
        data = [r.split("\t") for r in rows]
        if not data:
            self._set_status("รูปแบบข้อมูลว่าง")
            return
        col_id, row_id = self.tree.last_anchor
        if not row_id:
            kids = self.tree.get_children()
            if kids:
                row_id = kids[0]
            else:
                self.add_row()
                kids = self.tree.get_children()
                row_id = kids[0]
            col_id = "#1"
        start_row_index = list(self.tree.get_children()).index(row_id)
        start_col_index = int(col_id[1:]) - 1
        needed_rows = start_row_index + len(data) - len(self.tree.get_children())
        for _ in range(max(0, needed_rows)):
            self.add_row()
        all_iids = list(self.tree.get_children())
        for r, row_vals in enumerate(data):
            iid = all_iids[start_row_index + r]
            cur = list(self.tree.item(iid, "values"))
            for c, val in enumerate(row_vals):
                idx = start_col_index + c
                if 0 <= idx < len(COLUMNS) and COLUMNS[idx] != "Scan":
                    cur[idx] = val
            self.tree.item(iid, values=cur)
        self._set_status(f"วางข้อมูลจาก Excel {len(data)} แถวแล้ว")
        self.autosave()

    # ---------- Barcode Gen/Scan ----------
    def _random_code(self, length=10):
        alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
        return "".join(random.choice(alphabet) for _ in range(length))

    def generate_barcode_for_selected_or_empty(self):
        sel = self.tree.selection()
        count = 0
        targets = sel if sel else self.tree.get_children()
        for iid in targets:
            values = list(self.tree.item(iid, "values"))
            barcode_idx = COLUMNS.index("Barcode")
            if values[barcode_idx] and not sel:
                continue
            sku = values[COLUMNS.index("ชื่อSKU")] or values[COLUMNS.index("หมายเลขรายการ")]
            prefix = (sku or "ITEM").replace(" ", "").upper()
            code_core = datetime.datetime.now().strftime("%Y%m") + "-" + self._random_code(6)
            gen_code = f"{prefix}-{code_core}"
            values[barcode_idx] = gen_code
            self.tree.item(iid, values=values)
            count += 1
        self._set_status(f"สร้างบาร์โค้ด {count} รายการแล้ว")

    def generate_barcode_for_selected_or_empty_and_save(self):
        self.generate_barcode_for_selected_or_empty()
        self.autosave()

    def process_scan(self):
        data = self.scan_entry.get().strip()
        if not data:
            self._set_status("ไม่มีข้อมูลที่สแกน")
            self.scan_var.set("")
            self._maybe_focus_scan()
            return
        barcode_idx = COLUMNS.index("Barcode")
        scan_idx = COLUMNS.index("Scan")
        # ค้นหาแถวที่ barcode ตรงกับที่สแกน
        for iid in self.tree.get_children():
            values = list(self.tree.item(iid, "values"))
            if str(values[barcode_idx]).strip() == data:
                # ถ้าเคยสแกนแล้ว ให้เตือนและไม่เขียนทับ
                if str(values[scan_idx]).strip():
                    messagebox.showwarning("แจ้งเตือน", "สแกนไปแล้ว")
                    self._set_status(f"สแกนซ้ำ: {data}")
                    self._maybe_focus_scan()
                    return
                # ใส่เวลาที่มนุษย์อ่านได้แทนคำว่า pass
                ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                values[scan_idx] = ts
                self.tree.item(iid, values=values)
                self._set_status(f"สแกนสำเร็จ: {data} @ {ts}")
                self.scan_var.set("")
                self._maybe_focus_scan()
                return
        # ถ้าไม่พบบาร์โค้ดในตาราง
        messagebox.showwarning("ไม่พบ", f"ไม่พบบาร์โค้ด: {data}")
        self._set_status("ไม่พบบาร์โค้ดในตาราง")
        self.scan_var.set("")
        self._maybe_focus_scan()
        return

    def process_scan_and_save(self):
        self.process_scan()
        self.autosave()

    def _maybe_focus_scan(self):
        if self.auto_scan.get():
            self.scan_entry.focus_set()

    def recalc_all(self):
        n = 0
        for iid in self.tree.get_children():
            if self._recalc_row(iid):
                n += 1
        self._set_status(f"คำนวณต้นทุนทั้งหมด {n} แถวแล้ว")

    def recalc_all_and_save(self):
        self.recalc_all()
        self.autosave()

    def _recalc_row(self, iid):
        values = list(self.tree.item(iid, "values"))
        try:
            qty = float(values[COLUMNS.index("จำนวน (ชิ้น)")] or 0)
            price = float(values[COLUMNS.index("ราคาสินค้าต่อหน่วย (บาท)")] or 0)
            ship = float(values[COLUMNS.index("ค่าส่งรวม (บาท)")] or 0)
        except ValueError:
            return False
        total_cost = qty * price + ship
        unit_cost = total_cost / qty if qty else 0
        values[COLUMNS.index("ต้นทุนรวม (บาท)")] = f"{total_cost:.2f}"
        values[COLUMNS.index("ต้นทุนต่อตัว (บาท)")] = f"{unit_cost:.2f}"
        self.tree.item(iid, values=values)
        return True

    def clear_scan_selected(self):
        sel = self.tree.selection()
        if not sel:
            self._set_status("ยังไม่ได้เลือกแถว")
            return
        for iid in sel:
            values = list(self.tree.item(iid, "values"))
            values[COLUMNS.index("Scan")] = ""
            self.tree.item(iid, values=values)
        self._set_status(f"เคลียร์ Scan ให้ {len(sel)} แถวแล้ว")

    def clear_scan_selected_and_save(self):
        self.clear_scan_selected()
        self.autosave()

    # ---------- Barcode Export/Print ----------
    def _check_barcode_deps(self):
        try:
            import barcode  # type: ignore
            from barcode.writer import ImageWriter  # type: ignore
            return True
        except Exception:
            messagebox.showinfo(
                "ต้องติดตั้งไลบรารีเพิ่ม",
                "ฟังก์ชันพิมพ์บาร์โค้ดต้องติดตั้งแพ็กเกจเพิ่มเติมก่อน:\n\n"
                "python3 -m pip install python-barcode pillow\n\n"
                "ติดตั้งเสร็จแล้วเปิดโปรแกรมใหม่อีกครั้งครับ"
            )
            return False

    def _barcode_folder(self):
        folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "barcodes")
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception:
            pass
        return folder

    def export_barcodes_png(self, selected_only=True):
        if not self._check_barcode_deps():
            return
        import barcode  # type: ignore
        from barcode.writer import ImageWriter  # type: ignore

        barcode_idx = COLUMNS.index("Barcode")
        targets = self.tree.selection() if selected_only else self.tree.get_children()
        if not targets:
            self._set_status("ยังไม่ได้เลือกแถวสำหรับส่งออกบาร์โค้ด")
            return

        # Let user choose the output folder; cancel if none
        _default_dir = self._barcode_folder()
        init_dir = getattr(self, "settings", {}).get("last_export_dir", _default_dir)
        chosen_dir = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกบาร์โค้ด", initialdir=init_dir)
        if not chosen_dir:
            self._set_status("ยกเลิกการส่งออกบาร์โค้ด")
            return
        outdir = chosen_dir
        try:
            self.settings["last_export_dir"] = outdir
            self._save_settings()
        except Exception:
            pass
        count = 0
        for iid in targets:
            values = list(self.tree.item(iid, "values"))
            data = str(values[barcode_idx]).strip()
            if not data:
                continue
            try:
                code = barcode.get("code128", data, writer=ImageWriter())
                safe_name = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in data)
                filepath = os.path.join(outdir, f"{safe_name}.png")
                code.save(filepath, {"font_size": 12, "text_distance": 6, "module_height": 15})
                count += 1
            except Exception as e:
                print("Barcode error:", e)
                continue
        self._set_status(f"ส่งออกบาร์โค้ดเป็น PNG {count} ไฟล์ ไปที่โฟลเดอร์ 'barcodes' แล้ว")
        try:
            if sys.platform == "darwin":
                os.system(f"open '{outdir}'")
            elif os.name == "nt":
                os.startfile(outdir)  # type: ignore
            else:
                os.system(f"xdg-open '{outdir}'")
        except Exception:
            pass

    def print_selected_barcodes(self):
        outdir = self._barcode_folder()
        barcode_idx = COLUMNS.index("Barcode")
        targets = self.tree.selection()
        if not targets:
            self._set_status("ยังไม่ได้เลือกแถวสำหรับพิมพ์บาร์โค้ด")
            return
        files = []
        for iid in targets:
            values = list(self.tree.item(iid, "values"))
            data = str(values[barcode_idx]).strip()
            if not data:
                continue
            safe_name = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in data)
            png_path = os.path.join(outdir, f"{safe_name}.png")
            if os.path.exists(png_path):
                files.append(png_path)
        if not files:
            self._set_status("ยังไม่มีไฟล์ PNG สำหรับแถวที่เลือก กรุณา Export ก่อน")
            return
        opened = 0
        for fp in files:
            try:
                if sys.platform == "darwin":
                    os.system(f"open -a Preview '{fp}'")
                elif os.name == "nt":
                    os.startfile(fp, "print")  # type: ignore
                else:
                    os.system(f"xdg-open '{fp}'")
                opened += 1
            except Exception:
                continue
        self._set_status(f"เปิดไฟล์ {opened} รายการเพื่อพิมพ์แล้ว")


def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
