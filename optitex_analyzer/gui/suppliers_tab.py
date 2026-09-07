import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from . import theme


class SuppliersTabMixin:
    """Mixin לטאב ספקים: הוספה, מחיקה, הצגה ומסמכים."""

    def _create_suppliers_tab(self):
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="ספקים")
        tk.Label(tab, text="ניהול ספקים", font=(theme.FONT_FAMILY,16,'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=8)

        form = ttk.LabelFrame(tab, text="הוספת ספק", padding=10)
        form.pack(fill='x', padx=10, pady=6)

        self.sup_business_name_var = tk.StringVar()
        self.sup_first_name_var = tk.StringVar()
        self.sup_phone_var = tk.StringVar()
        self.sup_address_var = tk.StringVar()
        self.sup_business_var = tk.StringVar()
        self.sup_notes_var = tk.StringVar()

        labels = [
            ("שם עסק", self.sup_business_name_var, 18),
            ("שם פרטי", self.sup_first_name_var, 14),
            ("טלפון", self.sup_phone_var, 14),
            ("כתובת", self.sup_address_var, 25),
            ("מס' עסק", self.sup_business_var, 14),
            ("הערות", self.sup_notes_var, 25),
        ]
        for i,(lbl,var,w) in enumerate(labels):
            tk.Label(form, text=f"{lbl}:", font=(theme.FONT_FAMILY,10,'bold')).grid(row=0,column=i*2,sticky='w',padx=4,pady=4)
            tk.Entry(form, textvariable=var, width=w).grid(row=0,column=i*2+1,sticky='w',padx=2,pady=4)

        tk.Button(form, text="➕ הוסף", command=self._add_supplier_record, bg=theme.SUCCESS, fg='white').grid(row=0, column=len(labels)*2, padx=8)
        tk.Button(form, text="📄 מסמכים", command=self._open_selected_supplier_documents, bg=theme.PRIMARY, fg='white').grid(row=0, column=len(labels)*2+1, padx=4)
        tk.Button(form, text="🗑️ מחק נבחר", command=self._delete_selected_supplier, bg=theme.WARNING, fg='white').grid(row=0, column=len(labels)*2+2, padx=4)

        tree_frame = ttk.LabelFrame(tab, text="ספקים", padding=6)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('id','business_name','first_name','phone','address','business_number','notes','files','created_at')
        self.suppliers_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=14)
        headers = {
            'id':'ID','business_name':'שם עסק','first_name':'שם פרטי','phone':'טלפון','address':'כתובת',
            'business_number':'מספר עסק','notes':'הערות','files':'קבצים','created_at':'נוצר'
        }
        widths = {'id':45,'business_name':150,'first_name':110,'phone':110,'address':200,'business_number':110,'notes':160,'files':60,'created_at':140}
        for c in cols:
            self.suppliers_tree.heading(c, text=headers[c])
            self.suppliers_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.suppliers_tree.yview)
        self.suppliers_tree.configure(yscroll=vs.set)
        self.suppliers_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.suppliers_tree.bind('<Double-1>', lambda _e: self._open_selected_supplier_documents())
        self._load_suppliers_into_tree()

    def _supplier_files_count(self, rec):
        docs = rec.get('documents') if isinstance(rec, dict) else None
        return len(docs) if isinstance(docs, list) else 0

    def _load_suppliers_into_tree(self):
        if not hasattr(self, 'suppliers_tree'): return
        for item in self.suppliers_tree.get_children():
            self.suppliers_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'suppliers', []):
                self.suppliers_tree.insert('', 'end', values=(
                    rec.get('id'),
                    rec.get('business_name') or rec.get('name'),
                    rec.get('first_name',''),
                    rec.get('phone'),
                    rec.get('address'),
                    rec.get('business_number'),
                    rec.get('notes'),
                    self._supplier_files_count(rec),
                    rec.get('created_at')
                ))
        except Exception:
            pass

    def _add_supplier_record(self):
        business_name = self.sup_business_name_var.get().strip()
        if not business_name:
            messagebox.showerror("שגיאה", "חובה להזין שם עסק")
            return
        try:
            self.data_processor.add_supplier(
                business_name,
                self.sup_phone_var.get(),
                self.sup_address_var.get(),
                self.sup_business_var.get(),
                self.sup_notes_var.get(),
                self.sup_first_name_var.get(),
            )
            self._load_suppliers_into_tree()
            self.sup_business_name_var.set(''); self.sup_first_name_var.set(''); self.sup_phone_var.set(''); self.sup_address_var.set(''); self.sup_business_var.set(''); self.sup_notes_var.set('')
            # עדכן קומבואים בטאבים אחרים
            if hasattr(self, '_notify_suppliers_changed'):
                try: self._notify_suppliers_changed()
                except Exception: pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_supplier(self):
        sel = self.suppliers_tree.selection()
        if not sel:
            return
        ids = []
        for item in sel:
            vals = self.suppliers_tree.item(item, 'values')
            if vals:
                ids.append(int(vals[0]))
        deleted_any = False
        for _id in ids:
            if self.data_processor.delete_supplier(_id):
                deleted_any = True
        if deleted_any:
            self._load_suppliers_into_tree()
            if hasattr(self, '_notify_suppliers_changed'):
                try: self._notify_suppliers_changed()
                except Exception: pass

    def _selected_supplier_id(self):
        sel = self.suppliers_tree.selection()
        if not sel:
            return None
        vals = self.suppliers_tree.item(sel[0], 'values')
        if not vals:
            return None
        try:
            return int(vals[0])
        except (TypeError, ValueError):
            return None

    def _open_selected_supplier_documents(self):
        supplier_id = self._selected_supplier_id()
        if supplier_id is None:
            messagebox.showinfo("מסמכים", "יש לבחור ספק מהטבלה")
            return
        supplier = self.data_processor.get_supplier(supplier_id)
        if not supplier:
            messagebox.showerror("שגיאה", "ספק לא נמצא")
            return
        self._show_supplier_documents_dialog(supplier)

    def _show_supplier_documents_dialog(self, supplier):
        supplier_id = int(supplier.get('id'))
        display_name = supplier.get('business_name') or supplier.get('name') or f"ספק {supplier_id}"

        dialog = tk.Toplevel()
        dialog.title(f"מסמכי ספק — {display_name}")
        dialog.geometry("640x460")
        dialog.minsize(520, 360)
        dialog.transient(self.root if hasattr(self, 'root') else None)
        dialog.grab_set()
        dialog.configure(bg=theme.PAGE_BG)

        main = tk.Frame(dialog, bg=theme.PAGE_BG)
        main.pack(fill='both', expand=True, padx=16, pady=16)

        tk.Label(
            main,
            text=f"מסמכים של {display_name}",
            font=(theme.FONT_FAMILY, 13, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(anchor='w', pady=(0, 4))
        tk.Label(
            main,
            text="אפשר לייצא שיחה מוואטסאפ (ייצוא צ'אט, ללא מדיה) ולהעלות כאן קובץ txt או zip.",
            font=(theme.FONT_FAMILY, 9),
            bg=theme.PAGE_BG,
            fg=theme.SUBTEXT,
            wraplength=580,
            justify='right',
        ).pack(anchor='w', pady=(0, 10))

        tree_wrap = tk.Frame(main, bg=theme.PAGE_BG)
        tree_wrap.pack(fill='both', expand=True)
        cols = ('original_name', 'uploaded_at')
        docs_tree = ttk.Treeview(tree_wrap, columns=cols, show='headings', height=10)
        docs_tree.heading('original_name', text='שם קובץ')
        docs_tree.heading('uploaded_at', text='תאריך העלאה')
        docs_tree.column('original_name', width=380, anchor='w')
        docs_tree.column('uploaded_at', width=160, anchor='center')
        vs = ttk.Scrollbar(tree_wrap, orient='vertical', command=docs_tree.yview)
        docs_tree.configure(yscroll=vs.set)
        docs_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')

        def refresh_docs():
            for item in docs_tree.get_children():
                docs_tree.delete(item)
            current = self.data_processor.get_supplier(supplier_id) or {}
            for doc in current.get('documents') or []:
                docs_tree.insert('', 'end', iid=str(doc.get('id')), values=(
                    doc.get('original_name') or doc.get('filename') or '',
                    doc.get('uploaded_at') or '',
                ))
            self._load_suppliers_into_tree()

        def selected_doc_id():
            sel = docs_tree.selection()
            if not sel:
                return None
            try:
                return int(sel[0])
            except (TypeError, ValueError):
                return None

        def find_doc(doc_id):
            current = self.data_processor.get_supplier(supplier_id) or {}
            for doc in current.get('documents') or []:
                try:
                    if int(doc.get('id', 0)) == int(doc_id):
                        return doc
                except (TypeError, ValueError):
                    continue
            return None

        def upload_files():
            paths = filedialog.askopenfilenames(
                title="בחר קבצים לספק",
                filetypes=[
                    ("Excel", "*.xlsx *.xls"),
                    ("WhatsApp export", "*.txt *.zip"),
                    ("PDF", "*.pdf"),
                    ("Word", "*.doc *.docx"),
                    ("תמונות", "*.png *.jpg *.jpeg *.webp *.gif"),
                    ("כל הקבצים", "*.*"),
                ],
            )
            if not paths:
                return
            added = 0
            errors = []
            for path in paths:
                try:
                    self.data_processor.add_supplier_document(supplier_id, path)
                    added += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")
            refresh_docs()
            if errors:
                messagebox.showerror("שגיאה", "\n".join(errors), parent=dialog)
            elif added:
                messagebox.showinfo("מסמכים", f"הועלו {added} קבצים", parent=dialog)

        def open_selected():
            doc_id = selected_doc_id()
            if doc_id is None:
                messagebox.showinfo("מסמכים", "יש לבחור קובץ", parent=dialog)
                return
            doc = find_doc(doc_id)
            if not doc:
                messagebox.showerror("שגיאה", "מסמך לא נמצא", parent=dialog)
                return
            path = self.data_processor.supplier_document_path(supplier_id, doc.get('filename') or '')
            self._open_supplier_document_file(path)

        def delete_selected():
            doc_id = selected_doc_id()
            if doc_id is None:
                messagebox.showinfo("מסמכים", "יש לבחור קובץ", parent=dialog)
                return
            doc = find_doc(doc_id)
            label = (doc or {}).get('original_name') or (doc or {}).get('filename') or 'הקובץ'
            if not messagebox.askyesno("מחיקה", f"למחוק את {label}?", parent=dialog):
                return
            if self.data_processor.delete_supplier_document(supplier_id, doc_id):
                refresh_docs()
            else:
                messagebox.showerror("שגיאה", "לא ניתן למחוק את הקובץ", parent=dialog)

        docs_tree.bind('<Double-1>', lambda _e: open_selected())

        buttons = tk.Frame(main, bg=theme.PAGE_BG)
        buttons.pack(fill='x', pady=(12, 0))
        tk.Button(buttons, text="העלה קובץ", command=upload_files, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), width=14).pack(side='right', padx=4)
        tk.Button(buttons, text="פתח", command=open_selected, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), width=12).pack(side='right', padx=4)
        tk.Button(buttons, text="מחק קובץ", command=delete_selected, bg=theme.DANGER, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), width=12).pack(side='right', padx=4)
        tk.Button(buttons, text="סגור", command=dialog.destroy, bg=theme.DARK_2, fg='white', font=(theme.FONT_FAMILY, 10), width=10).pack(side='left', padx=4)

        refresh_docs()

    def _open_supplier_document_file(self, file_path):
        try:
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                messagebox.showerror("שגיאה", f"קובץ לא נמצא: {file_path}")
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן לפתוח את הקובץ: {str(e)}")
