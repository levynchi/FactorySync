import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import re  # עבור פיצול מידות מרובות (וריאנטים)

class ProductsCatalogTabMixin:
    """Mixin לטאב ניהול קטלוג מוצרים (הוספה / מחיקה / ייצוא)."""
    def _create_products_catalog_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="קטלוג מוצרים")
        tk.Label(tab, text="ניהול קטלוג מוצרים", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)
        form = ttk.LabelFrame(tab, text="הוספת מוצר", padding=10)
        form.pack(fill='x', padx=10, pady=6)
        self.prod_name_var = tk.StringVar(); self.prod_size_var = tk.StringVar(); self.prod_fabric_type_var = tk.StringVar(); self.prod_fabric_color_var = tk.StringVar(); self.prod_print_name_var = tk.StringVar()
        # עדכון: שדה המידה תומך במספר וריאנטים בבת אחת מופרדים בפסיק / רווח (למשל: "0-3,3-6,6-12")
        # עדכון: כל אחד מהשדות (מידה / סוג בד / צבע בד / שם פרינט) תומך ברשימת ערכים מופרדים בפסיק / רווחים ליצירת וריאנטים מרובים.
        labels = [
            ("שם מוצר", self.prod_name_var, 25),
            ("מידות (פסיק)", self.prod_size_var, 18),
            ("סוגי בד (פסיק)", self.prod_fabric_type_var, 18),
            ("צבעי בד (פסיק)", self.prod_fabric_color_var, 18),
            ("שמות פרינט (פסיק)", self.prod_print_name_var, 18)
        ]
        for i,(lbl,var,width) in enumerate(labels):
            tk.Label(form, text=f"{lbl}:", font=('Arial',10,'bold')).grid(row=0, column=i*2, sticky='w', padx=4, pady=4)
            tk.Entry(form, textvariable=var, width=width).grid(row=0, column=i*2+1, sticky='w', padx=2, pady=4)
        tk.Button(form, text="➕ הוסף", command=self._add_product_catalog_entry, bg='#27ae60', fg='white').grid(row=0, column=len(labels)*2, padx=8)
        tk.Button(form, text="🗑️ מחק נבחר", command=self._delete_selected_product_entry, bg='#e67e22', fg='white').grid(row=0, column=len(labels)*2+1, padx=4)
        tk.Button(form, text="💾 ייצוא ל-Excel", command=self._export_products_catalog, bg='#2c3e50', fg='white').grid(row=0, column=len(labels)*2+2, padx=4)
        # Treeview
        tree_frame = ttk.LabelFrame(tab, text="מוצרים", padding=6)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('id','name','size','fabric_type','fabric_color','print_name','created_at')
        self.products_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        headers = {'id':'ID','name':'שם מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'שם פרינט','created_at':'נוצר'}
        widths = {'id':50,'name':180,'size':80,'fabric_type':120,'fabric_color':120,'print_name':140,'created_at':140}
        for c in cols:
            self.products_tree.heading(c, text=headers[c])
            self.products_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.products_tree.yview)
        self.products_tree.configure(yscroll=vs.set)
        self.products_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self._load_products_catalog_into_tree()

    # Loading
    def _load_products_catalog_into_tree(self):
        if not hasattr(self, 'products_tree'): return
        for item in self.products_tree.get_children(): self.products_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'products_catalog', []):
                self.products_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('size'), rec.get('fabric_type'), rec.get('fabric_color'), rec.get('print_name'), rec.get('created_at')))
        except Exception:
            pass

    def _add_product_catalog_entry(self):
        """הוספת מוצר/ים לקטלוג.

        תמיכה בוריאנטים מרובים בכל אחד מהשדות: מידה / סוג בד / צבע בד / שם פרינט.
        כל שדה יכול להכיל ערכים מופרדים בפסיקים / רווחים / נקודה-פסיק.
        נוצרת כל הצירוף (Cartesian product) של הערכים שסופקו בכל השדות (למעט שם המוצר).
        מניעת כפילויות: אם רשומה באותו שם + מידה + סוג בד + צבע בד + שם פרינט כבר קיימת – נדלג.
        """
        name = self.prod_name_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם מוצר")
            return
        sizes_raw = self.prod_size_var.get().strip()
        ftypes_raw = self.prod_fabric_type_var.get().strip()
        fcolors_raw = self.prod_fabric_color_var.get().strip()
        prints_raw = self.prod_print_name_var.get().strip()

        def _split(raw):
            if not raw:
                return ['']  # ערך יחיד ריק (כדי לא לאבד רשומה)
            return [s.strip() for s in re.split(r'[;,.\s]+', raw) if s.strip()]

        size_tokens = _split(sizes_raw)
        ft_tokens = _split(ftypes_raw)
        fc_tokens = _split(fcolors_raw)
        pn_tokens = _split(prints_raw)

        # נרמול טוקנים של מידות – תיקון שגיאות הקלדה נפוצות כמו "-1218" או "1218" -> "12-18"
        def _normalize_size(tok: str) -> str:
            t = tok.strip()
            if not t:
                return t
            # הסר מקף מוביל שגוי
            if t.startswith('-'):
                t = t[1:]
            # אם אין מקף וקיימות רק ספרות באורך 3-4 → ננסה לפצל באמצע (לדוגמה 1218 -> 12-18)
            if '-' not in t and t.isdigit() and 3 <= len(t) <= 4:
                # ננסה לחלק לשתי קבוצות (חצי ראשון / חצי שני)
                mid = len(t)//2
                a, b = t[:mid], t[mid:]
                # ודא ששני החלקים מספריים וקטנים מ-60 (חודשים סבירים)
                try:
                    ai = int(a); bi = int(b)
                    if 0 <= ai <= 60 and 0 <= bi <= 60 and ai < bi:
                        return f"{ai}-{bi}"
                except Exception:
                    pass
            return t

        size_tokens = [_normalize_size(s) for s in size_tokens]

        from itertools import product
        combos = list(product(size_tokens, ft_tokens, fc_tokens, pn_tokens))
        if not combos:
            combos = [( '', '', '', '' )]

        # סט לרשומות קיימות למניעת כפילות
        existing = set()
        try:
            for rec in getattr(self.data_processor, 'products_catalog', []):
                existing.add((rec.get('name','').strip(), rec.get('size','').strip(), rec.get('fabric_type','').strip(), rec.get('fabric_color','').strip(), rec.get('print_name','').strip()))
        except Exception:
            existing = set()

        # אם המשתמש הזין וריאנט יחיד (צירוף יחיד) ונמצא שהוא כבר קיים – נחסום ונודיע במפורש
        if len(combos) == 1:
            only_sz, only_ft, only_fc, only_pn = combos[0]
            single_key = (name, only_sz, only_ft, only_fc, only_pn)
            if single_key in existing:
                messagebox.showinfo(
                    "כפילות",
                    "המוצר עם הנתונים הללו כבר קיים במערכת:\n"
                    f"שם: {name}\nמידה: {only_sz or '-'}\nסוג בד: {only_ft or '-'}\nצבע בד: {only_fc or '-'}\nשם פרינט: {only_pn or '-'}"
                )
                return

        added = 0
        try:
            for sz, ft, fc, pn in combos:
                key = (name, sz, ft, fc, pn)
                if key in existing:
                    continue
                new_id = self.data_processor.add_product_catalog_entry(name, sz, ft, fc, pn)
                existing.add(key)
                added += 1
                self.products_tree.insert('', 'end', values=(new_id, name, sz, ft, fc, pn, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.prod_name_var.set(''); self.prod_size_var.set(''); self.prod_fabric_type_var.set(''); self.prod_fabric_color_var.set(''); self.prod_print_name_var.set('')
            if added > 1:
                messagebox.showinfo("הצלחה", f"נוספו {added} וריאנטים למוצר '{name}'")
            elif added == 1:
                # שקט – הוספה יחידה
                pass
            else:
                # כל הצירופים הוזנו כבר בעבר
                messagebox.showinfo("כפילות", "כל הצירופים שהוזנו כבר קיימים – לא נוספו מוצרים חדשים")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_product_entry(self):
        sel = self.products_tree.selection()
        if not sel: return
        ids = []
        for item in sel:
            vals = self.products_tree.item(item, 'values')
            if vals:
                ids.append(int(vals[0]))
        if not ids: return
        deleted_any = False
        for _id in ids:
            if self.data_processor.delete_product_catalog_entry(_id):
                deleted_any = True
        if deleted_any:
            self._load_products_catalog_into_tree()

    def _export_products_catalog(self):
        if not getattr(self.data_processor, 'products_catalog', []):
            messagebox.showerror("שגיאה", "אין מוצרים לייצוא")
            return
        file_path = filedialog.asksaveasfilename(title="ייצוא קטלוג מוצרים", defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
        if not file_path: return
        try:
            self.data_processor.export_products_catalog_to_excel(file_path)
            messagebox.showinfo("הצלחה", "הקטלוג יוצא בהצלחה")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
