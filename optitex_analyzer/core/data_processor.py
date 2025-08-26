"""
עיבוד נתונים וייצוא
"""

import pandas as pd
import json
import os
from datetime import datetime
from typing import Dict, List, Any

class DataProcessor:
	"""מעבד נתונים וייצוא"""
    
	def __init__(self, drawings_file: str = "drawings_data.json", fabrics_inventory_file: str = "fabrics_inventory.json", fabrics_imports_file: str = "fabrics_import_logs.json", supplier_receipts_file: str = "supplier_receipts.json", products_catalog_file: str = "products_catalog.json", suppliers_file: str = "suppliers.json", supplier_intakes_file: str = "supplier_intakes.json", delivery_notes_file: str = "delivery_notes.json", sewing_accessories_file: str = "sewing_accessories.json", categories_file: str = "categories.json", product_sizes_file: str = "product_sizes.json", fabric_types_file: str = "fabric_types.json", fabric_colors_file: str = "fabric_colors.json", print_names_file: str = "print_names.json", fabric_categories_file: str = "fabric_categories.json", model_names_file: str = "model_names.json", main_categories_file: str = "main_categories.json", fabrics_intakes_file: str = "fabrics_intakes.json"):
		"""
		שימו לב: בעבר השתמשנו בקובץ אחד (supplier_receipts.json) עבור שני סוגי הרשומות
		(supplier_intake / delivery_note). כעת הם מופרדים לשני קבצים: supplier_intakes.json ו‑delivery_notes.json.
		עדיין נשמרת תאימות לאחור: אם הקובץ הישן קיים – מתבצעת העברה חד‑פעמית.
		"""
		self.drawings_file = drawings_file
		self.fabrics_inventory_file = fabrics_inventory_file
		self.fabrics_imports_file = fabrics_imports_file
		# legacy combined file (לצורך מיגרציה בלבד)
		self.supplier_receipts_file = supplier_receipts_file
		# new split files
		self.supplier_intakes_file = supplier_intakes_file
		self.delivery_notes_file = delivery_notes_file
		# קובץ תיעוד 'תעודת קליטת בדים'
		self.fabrics_intakes_file = fabrics_intakes_file
		self.products_catalog_file = products_catalog_file
		self.sewing_accessories_file = sewing_accessories_file
		self.categories_file = categories_file
		self.suppliers_file = suppliers_file
		# קבצי תכונות מוצר
		self.product_sizes_file = product_sizes_file
		self.fabric_types_file = fabric_types_file
		self.fabric_colors_file = fabric_colors_file
		self.print_names_file = print_names_file
		self.fabric_categories_file = fabric_categories_file
		# קובץ 'שם הדגם'
		self.model_names_file = model_names_file
		# קובץ קטגוריה ראשית
		self.main_categories_file = main_categories_file
		# load base datasets
		self.drawings_data = self.load_drawings_data()
		self.fabrics_inventory = self.load_fabrics_inventory()
		self.fabrics_import_logs = self.load_fabrics_import_logs()
		self.products_catalog = self.load_products_catalog()
		self.sewing_accessories = self.load_sewing_accessories()
		self.categories = self.load_categories()
		self.main_categories = self.load_main_categories()
		# טעינת תכונות מוצר (מידות, סוגי בד, צבעי בד, שמות פרינט)
		self.product_sizes = self.load_product_sizes()
		self.product_fabric_types = self.load_fabric_types()
		self.product_fabric_colors = self.load_fabric_colors()
		self.product_print_names = self.load_print_names()
		self.product_fabric_categories = self.load_fabric_categories()
		self.product_model_names = self.load_model_names()
		# Ensure all attribute records have a main_category (default 'בגדים')
		try:
			self._ensure_main_category_on_attributes()
		except Exception:
			pass
		self.suppliers = self.load_suppliers()
		# load split receipts (may be empty on first run)
		self.supplier_intakes = self._load_json_list(self.supplier_intakes_file)
		self.delivery_notes = self._load_json_list(self.delivery_notes_file)
		# טעינת תעודות קליטת בדים
		self.fabrics_intakes = self._load_json_list(self.fabrics_intakes_file)
		# migration from legacy combined file if needed
		self._migrate_legacy_supplier_receipts()
		# combined view (backward compatibility for old UI code)
		self.supplier_receipts = self.supplier_intakes + self.delivery_notes

	# ===== Fabrics Intake Receipts (תעודת קליטת בדים) =====
	def add_fabrics_intake(self, barcodes: List[str], packages: List[Dict] | None = None, *, supplier: str = '', date_str: str = '') -> int:
		"""יוצר תיעוד חדש של 'תעודת קליטת בדים' ושומר לקובץ. מחזיר ID חדש."""
		try:
			if not barcodes:
				raise ValueError("אין ברקודים לקליטה")
			new_id = self._next_id(getattr(self, 'fabrics_intakes', []) or [])
			date_v = (date_str or datetime.now().strftime('%Y-%m-%d')).strip()
			record = {
				'id': new_id,
				'date': date_v,
				'supplier': (supplier or '').strip(),
				'barcodes': list(barcodes),
				'count_barcodes': len(barcodes),
				'packages': packages or [],
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.fabrics_intakes.append(record)
			self._save_json_list(self.fabrics_intakes_file, self.fabrics_intakes)
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת תעודת קליטת בדים: {e}")

	def delete_fabrics_intake(self, rec_id: int) -> bool:
		"""מחיקת תעודת קליטת בדים לפי ID. מחזיר True אם נמחקה."""
		before = len(getattr(self, 'fabrics_intakes', []) or [])
		try:
			self.fabrics_intakes = [r for r in self.fabrics_intakes if int(r.get('id', -1)) != int(rec_id)]
			if len(self.fabrics_intakes) != before:
				self._save_json_list(self.fabrics_intakes_file, self.fabrics_intakes)
				return True
			return False
		except Exception:
			self.fabrics_intakes = [r for r in self.fabrics_intakes if (r.get('id') != rec_id)]
			if len(self.fabrics_intakes) != before:
				self._save_json_list(self.fabrics_intakes_file, self.fabrics_intakes)
				return True
			return False

	def refresh_fabrics_intakes(self):
		"""רענון תעודות קליטת בדים מהדיסק"""
		self.fabrics_intakes = self._load_json_list(self.fabrics_intakes_file)

	def load_suppliers(self) -> List[Dict]:
		"""טעינת רשימת ספקים"""
		try:
			if os.path.exists(self.suppliers_file):
				with open(self.suppliers_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת ספקים: {e}"); return []

	def save_suppliers(self) -> bool:
		"""שמירת ספקים"""
		try:
			with open(self.suppliers_file, 'w', encoding='utf-8') as f:
				json.dump(self.suppliers, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת ספקים: {e}"); return False

	def add_supplier(self, business_name: str, phone: str = '', address: str = '', business_number: str = '', notes: str = '', first_name: str = '') -> int:
		"""הוספת ספק חדש (שם עסק + שם פרטי אופציונלי). מחזיר ID חדש.

		שדות:
		- business_name: שם העסק (חובה)
		- first_name: שם פרטי איש קשר (לא חובה)
		- phone, address, business_number, notes: פרטים נוספים.
		"""
		try:
			if not business_name:
				raise ValueError("חובה להזין שם עסק")
			new_id = max([s.get('id',0) for s in self.suppliers], default=0) + 1
			record = {
				'id': new_id,
				'business_name': business_name.strip(),
				'first_name': first_name.strip(),
				'phone': phone.strip(),
				'address': address.strip(),
				'business_number': business_number.strip(),
				'notes': notes.strip(),
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.suppliers.append(record)
			self.save_suppliers()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת ספק: {e}")

	def delete_supplier(self, supplier_id: int) -> bool:
		"""מחיקת ספק לפי ID"""
		before = len(self.suppliers)
		self.suppliers = [s for s in self.suppliers if int(s.get('id',0)) != int(supplier_id)]
		if len(self.suppliers) != before:
			self.save_suppliers(); return True
		return False
    
	def load_drawings_data(self) -> List[Dict]:
		"""טעינת נתוני ציורים מקומיים"""
		try:
			if os.path.exists(self.drawings_file):
				with open(self.drawings_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת נתוני ציורים: {e}")
			return []
    
	def save_drawings_data(self) -> bool:
		"""שמירת נתוני ציורים"""
		try:
			with open(self.drawings_file, 'w', encoding='utf-8') as f:
				json.dump(self.drawings_data, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת נתוני ציורים: {e}")
			return False
    
	# (הוסר) טיפול ב"ציורים חוזרים" – המערכת לא שומרת יותר רשומות כאלה נפרדות.

	# ===== Supplier Receipts (Manual Products Intake) =====
	def _load_json_list(self, path: str) -> List[Dict]:
		try:
			if os.path.exists(path):
				with open(path, 'r', encoding='utf-8') as f:
					return json.load(f) or []
			return []
		except Exception:
			return []

	def _save_json_list(self, path: str, data: List[Dict]) -> bool:
		try:
			with open(path, 'w', encoding='utf-8') as f:
				json.dump(data, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קובץ {path}: {e}"); return False

	def _migrate_legacy_supplier_receipts(self):
		"""אם הקובץ הישן קיים ויש בו רשומות – מפצל אותן לשני הקבצים החדשים.
		המיגרציה מתבצעת רק אם אחד מהקבצים החדשים ריק.
		"""
		try:
			if not os.path.exists(self.supplier_receipts_file):
				return
			legacy_list = self._load_json_list(self.supplier_receipts_file)
			if not legacy_list:
				return
			# אם כבר קיימים נתונים בקבצים החדשים – נניח שכבר בוצעה מיגרציה
			if self.supplier_intakes or self.delivery_notes:
				return
			for rec in legacy_list:
				kind = rec.get('receipt_kind') or 'supplier_intake'
				if kind == 'delivery_note':
					self.delivery_notes.append(rec)
				else:
					self.supplier_intakes.append(rec)
			# שמירה לקבצים החדשים
			self._save_json_list(self.supplier_intakes_file, self.supplier_intakes)
			self._save_json_list(self.delivery_notes_file, self.delivery_notes)
			# אפשר לשמר גיבוי של הקובץ הישן בשם אחר (לא מוחקים כדי לאבד היסטוריה)
		except Exception as e:
			print(f"שגיאה במיגרציית קליטות ספק: {e}")

	def add_supplier_receipt(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, receipt_kind: str = "supplier_intake", *, arrival_date: str = "", supplier_doc_number: str = "") -> int:
		"""שכבת תאימות – מפנה לפונקציה המתאימה לפי receipt_kind."""
		if receipt_kind == 'delivery_note':
			return self.add_delivery_note(supplier, date_str, lines, packages, arrival_date=arrival_date, supplier_doc_number=supplier_doc_number)
		return self.add_supplier_intake(supplier, date_str, lines, packages, arrival_date=arrival_date, supplier_doc_number=supplier_doc_number)

	def _next_id(self, records: List[Dict]) -> int:
		return max([r.get('id', 0) for r in records], default=0) + 1

	def add_supplier_intake(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, *, arrival_date: str = "", supplier_doc_number: str = "") -> int:
		try:
			if not supplier: raise ValueError("חסר שם ספק")
			if not lines: raise ValueError("אין שורות לקליטה")
			new_id = self._next_id(self.supplier_intakes)
			total_quantity = sum(int(l.get('quantity',0)) for l in lines)
			record = {
				'id': new_id,
				'supplier': supplier,
				'date': date_str,
				'lines': lines,
				'total_quantity': total_quantity,
				'packages': packages or [],
				'receipt_kind': 'supplier_intake',
				'arrival_date': arrival_date or "",
				'supplier_doc_number': supplier_doc_number or "",
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.supplier_intakes.append(record)
			self._save_json_list(self.supplier_intakes_file, self.supplier_intakes)
			self._rebuild_combined_receipts()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קליטת ספק: {e}")

	def add_delivery_note(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, *, arrival_date: str = "", supplier_doc_number: str = "") -> int:
		try:
			if not supplier: raise ValueError("חסר שם ספק")
			if not lines: raise ValueError("אין שורות לקליטה")
			new_id = self._next_id(self.delivery_notes)
			total_quantity = sum(int(l.get('quantity',0)) for l in lines)
			record = {
				'id': new_id,
				'supplier': supplier,
				'date': date_str,
				'lines': lines,
				'total_quantity': total_quantity,
				'packages': packages or [],
				'receipt_kind': 'delivery_note',
				'arrival_date': arrival_date or "",
				'supplier_doc_number': supplier_doc_number or "",
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.delivery_notes.append(record)
			self._save_json_list(self.delivery_notes_file, self.delivery_notes)
			self._rebuild_combined_receipts()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת תעודת משלוח: {e}")

	def _rebuild_combined_receipts(self):
		self.supplier_receipts = self.supplier_intakes + self.delivery_notes

	def delete_supplier_intake(self, receipt_id: int) -> bool:
		"""מוחק קליטת ספק (supplier_intake) לפי ID. מחזיר True אם נמחקה רשומה."""
		before = len(self.supplier_intakes)
		try:
			self.supplier_intakes = [r for r in self.supplier_intakes if int(r.get('id', -1)) != int(receipt_id)]
		except Exception:
			self.supplier_intakes = [r for r in self.supplier_intakes if (r.get('id') != receipt_id)]
		if len(self.supplier_intakes) != before:
			self._save_json_list(self.supplier_intakes_file, self.supplier_intakes)
			self._rebuild_combined_receipts()
			return True
		return False

	def delete_delivery_note(self, note_id: int) -> bool:
		"""מוחק תעודת משלוח (delivery_note) לפי ID. מחזיר True אם נמחקה רשומה."""
		before = len(self.delivery_notes)
		try:
			self.delivery_notes = [r for r in self.delivery_notes if int(r.get('id', -1)) != int(note_id)]
		except Exception:
			self.delivery_notes = [r for r in self.delivery_notes if (r.get('id') != note_id)]
		if len(self.delivery_notes) != before:
			self._save_json_list(self.delivery_notes_file, self.delivery_notes)
			self._rebuild_combined_receipts()
			return True
		return False

	def get_delivery_notes(self) -> List[Dict]:
		return list(self.delivery_notes)

	def refresh_supplier_receipts(self):
		self.supplier_intakes = self._load_json_list(self.supplier_intakes_file)
		self.delivery_notes = self._load_json_list(self.delivery_notes_file)
		self._rebuild_combined_receipts()
    
	# (הוסר) פעולות על ציורים חוזרים – get/refresh/add/delete
    
	def results_to_dataframe(self, results: List[Dict]) -> pd.DataFrame:
		"""המרת תוצאות ל-DataFrame"""
		if not results:
			return pd.DataFrame()
		return pd.DataFrame(results)
    
	def export_to_excel(self, results: List[Dict], file_path: str) -> bool:
		"""ייצוא תוצאות ל-Excel"""
		try:
			df = self.results_to_dataframe(results)
			if df.empty:
				raise ValueError("אין נתונים לייצוא")
            
			df.to_excel(file_path, index=False)
			return True
		except Exception as e:
			raise Exception(f"שגיאה בייצוא ל-Excel: {str(e)}")
    
	def add_to_local_table(self, results: List[Dict], file_name: str = "", fabric_type: str = "", recipient_supplier: str = "") -> int:
		"""הוספה לטבלה המקומית"""
		try:
			# יצירת ID חדש
			new_id = max([r.get('id', 0) for r in self.drawings_data], default=0) + 1
            
			# יצירת רשומה חדשה
			record = {
				'id': new_id,
				'שם הקובץ': os.path.splitext(os.path.basename(file_name))[0] if file_name else 'לא ידוע',
				'תאריך יצירה': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
				'מוצרים': [],
				'סך כמויות': 0,
				# סטטוס ציור: נשלח (ברירת מחדל) | הוחזר | טרם נשלח
				'status': 'נשלח'
			}
			if fabric_type:
				record['סוג בד'] = fabric_type
			# שמירת יעד/ספק אליו נשלח הציור (אם נבחר בממיר)
			recipient_supplier = (recipient_supplier or '').strip()
			if recipient_supplier:
				record['נמען'] = recipient_supplier
				# עמודה לוגית לנוחות סינון עתידי
				record['נשלח לספק'] = True
            
			# קיבוץ לפי מוצרים
			df = pd.DataFrame(results)
			for product_name in df['שם המוצר'].unique():
				product_data = df[df['שם המוצר'] == product_name]
                
				product_info = {
					'שם המוצר': product_name,
					'מידות': []
				}
                
				for _, row in product_data.iterrows():
					size_info = {
						'מידה': row['מידה'],
						'כמות': row['כמות'],
						'הערה': row['הערה']
					}
					product_info['מידות'].append(size_info)
					record['סך כמויות'] += row['כמות']
                
				record['מוצרים'].append(product_info)
            
			# הוספה לרשימה ושמירה
			self.drawings_data.append(record)
			self.save_drawings_data()
            
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספה לטבלה המקומית: {str(e)}")
    
	def delete_drawing(self, drawing_id: int) -> bool:
		"""מחיקת ציור לפי ID"""
		try:
			# השוואה עמידה לסוגים (int/str)
			self.drawings_data = [r for r in self.drawings_data if str(r.get('id')) != str(drawing_id)]
			return self.save_drawings_data()
		except Exception as e:
			print(f"שגיאה במחיקת ציור: {e}")
			return False
    
	def clear_all_drawings(self) -> bool:
		"""מחיקת כל הציורים"""
		try:
			self.drawings_data = []
			return self.save_drawings_data()
		except Exception as e:
			print(f"שגיאה במחיקת כל הציורים: {e}")
			return False
    
	def export_drawings_to_excel(self, file_path: str) -> bool:
		"""ייצוא כל הציורים ל-Excel"""
		try:
			if not self.drawings_data:
				raise ValueError("אין ציורים לייצוא")
            
			rows = []
			for record in self.drawings_data:
				for product in record.get('מוצרים', []):
					for size_info in product.get('מידות', []):
						rows.append({
							'ID רשומה': record.get('id', ''),
							'שם הקובץ': record.get('שם הקובץ', ''),
							'תאריך יצירה': record.get('תאריך יצירה', ''),
							'שם המוצר': product.get('שם המוצר', ''),
							'מידה': size_info.get('מידה', ''),
							'כמות': size_info.get('כמות', 0),
							'הערה': size_info.get('הערה', '')
						})
            
			if not rows:
				raise ValueError("אין נתונים לייצוא")
            
			df = pd.DataFrame(rows)
			df.to_excel(file_path, index=False)
			return True
            
		except Exception as e:
			raise Exception(f"שגיאה בייצוא ציורים: {str(e)}")
    
	def get_drawing_by_id(self, drawing_id: int) -> Dict:
		"""קבלת ציור לפי ID"""
		for record in self.drawings_data:
			if record.get('id') == drawing_id:
				return record
		return {}
    
	def refresh_drawings_data(self):
		"""רענון נתוני הציורים מהקובץ"""
		self.drawings_data = self.load_drawings_data()

	def update_drawing_status(self, drawing_id: int, new_status: str) -> bool:
		"""עדכון סטטוס לציור (נשלח / הוחזר / טרם נשלח). מחזיר True אם עודכן."""
		try:
			changed = False
			for rec in self.drawings_data:
				if rec.get('id') == drawing_id:
					if rec.get('status') != new_status:
						rec['status'] = new_status
						changed = True
						break
			if changed:
				self.save_drawings_data()
			return changed
		except Exception as e:
			print(f"שגיאה בעדכון סטטוס ציור: {e}")
			return False

	def update_drawing_layers(self, drawing_id: int, layers: int) -> bool:
		"""עדכון ערך 'שכבות' לציור. מחזיר True אם בוצע שינוי ונשמר."""
		try:
			if layers is None:
				return False
			changed = False
			for rec in self.drawings_data:
				if rec.get('id') == drawing_id:
					prev = rec.get('שכבות')
					try:
						prev_i = int(prev) if prev is not None else None
					except Exception:
						prev_i = None
					if prev_i != int(layers):
						rec['שכבות'] = int(layers)
						changed = True
					break
			if changed:
				self.save_drawings_data()
			return changed
		except Exception as e:
			print(f"שגיאה בעדכון שכבות לציור: {e}")
			return False

	# ===== Fabrics Inventory =====
	def load_fabrics_inventory(self) -> List[Dict]:
		"""טעינת מלאי בדים"""
		try:
			if os.path.exists(self.fabrics_inventory_file):
				with open(self.fabrics_inventory_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת מלאי בדים: {e}")
			return []

	def save_fabrics_inventory(self) -> bool:
		"""שמירת מלאי בדים"""
		try:
			with open(self.fabrics_inventory_file, 'w', encoding='utf-8') as f:
				json.dump(self.fabrics_inventory, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת מלאי בדים: {e}")
			return False

	def import_fabrics_csv(self, file_path: str) -> int:
		"""ייבוא משלוח בדים מקובץ CSV
		:return: מספר רשומות שנוספו
		"""
		import csv
		added = 0
		try:
			# ניצור מראש לוג (לשייך ID לרשומות) – נעדכן כמות לאחר הספירה
			temp_log_id = self.add_fabric_import_log(file_path, 0)
			# ניסיון קריאה ב-UTF-8 עם BOM, ואם לא מצליח נופל ל-latin-1
			encodings = ['utf-8-sig', 'cp1255', 'latin-1']
			rows = []
			for enc in encodings:
				try:
					with open(file_path, 'r', encoding=enc, newline='') as f:
						reader = csv.DictReader(f)
						for r in reader:
							rows.append(r)
					if rows:
						break
				except Exception:
					rows = []
					continue
			if not rows:
				raise ValueError("לא נקראו רשומות מהקובץ")

			# ניקוי כותרות וטרנספורמציה
			def to_float(val):
				try:
					if val is None:
						return 0.0
					v = str(val).strip().replace(',', '')
					if v == '':
						return 0.0
					return float(v)
				except Exception:
					return 0.0

			for r in rows:
				cleaned = { (k.strip() if k else ''): (v.strip() if isinstance(v, str) else v) for k,v in r.items() }
				record = {
					'barcode': cleaned.get('BARCODE NO', ''),
					'fabric_type': cleaned.get('סוג בד', ''),
					'color_name': cleaned.get('COLOR NAME', ''),
					'color_no': cleaned.get('COLOR NO', ''),
					'design_code': cleaned.get('Desen Kodu', ''),
					'width': cleaned.get('WIDTH', ''),
					'gr': cleaned.get('GR', ''),
					'net_kg': to_float(cleaned.get('NET KG')),
					'gross_kg': to_float(cleaned.get('GROSS KG')),
					'meters': to_float(cleaned.get('METER')),
					'price': to_float(cleaned.get('PRICE')) if 'PRICE' in cleaned else to_float(cleaned.get('PRICE ','0')),
					'total': to_float(cleaned.get('TOTAL')) if 'TOTAL' in cleaned else to_float(cleaned.get('TOTAL  ','0')),
					'location': cleaned.get('location', ''),
					'last_modified': cleaned.get('Last Modified', ''),
					'purpose': cleaned.get('מטרה', ''),
					'import_log_id': temp_log_id,
					# סטטוס מלאי: ברירת מחדל "במלאי" בעת קליטה
					'status': 'במלאי'
				}
				if record['barcode']:
					self.fabrics_inventory.append(record)
					added += 1

			if added:
				self.save_fabrics_inventory()
				# עדכון הלוג עם מספר רשומות
				for log in self.fabrics_import_logs:
					if log.get('id') == temp_log_id:
						log['records_added'] = added
						break
				self.save_fabrics_import_logs()
			else:
				# אם לא נוספו רשומות – למחוק לוג ריק
				self.fabrics_import_logs = [l for l in self.fabrics_import_logs if l.get('id') != temp_log_id]
				self.save_fabrics_import_logs()
			return added
		except Exception as e:
			raise Exception(f"שגיאה בייבוא מלאי בדים: {str(e)}")

	def get_fabrics_summary(self) -> Dict[str, Any]:
		"""סיכום מלאי בדים"""
		total = len(self.fabrics_inventory)
		total_meters = sum(item.get('meters', 0) for item in self.fabrics_inventory)
		total_net = sum(item.get('net_kg', 0) for item in self.fabrics_inventory)
		return {
			'total_records': total,
			'total_meters': total_meters,
			'total_net_kg': total_net
		}

	def refresh_fabrics_inventory(self):
		"""רענון מלאי בדים מהדיסק"""
		self.fabrics_inventory = self.load_fabrics_inventory()

	def export_fabrics_to_excel(self, file_path: str) -> bool:
		"""ייצוא מלאי הבדים לקובץ Excel"""
		try:
			if not self.fabrics_inventory:
				raise ValueError("אין נתוני מלאי לייצוא")
			df = pd.DataFrame(self.fabrics_inventory)
			df.to_excel(file_path, index=False)
			return True
		except Exception as e:
			raise Exception(f"שגיאה בייצוא מלאי בדים: {str(e)}")

	# ===== Fabrics Import Logs =====
	def load_fabrics_import_logs(self) -> List[Dict]:
		"""טעינת לוג ייבוא קבצי מלאי בדים"""
		try:
			if os.path.exists(self.fabrics_imports_file):
				with open(self.fabrics_imports_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת לוג ייבוא בדים: {e}")
			return []

	def save_fabrics_import_logs(self) -> bool:
		"""שמירת לוג ייבוא קבצי מלאי בדים"""
		try:
			with open(self.fabrics_imports_file, 'w', encoding='utf-8') as f:
				json.dump(self.fabrics_import_logs, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת לוג ייבוא בדים: {e}")
			return False

	def add_fabric_import_log(self, file_path: str, records_added: int):
		"""הוספת רשומת לוג חדשה עבור קובץ שיובא"""
		try:
			log_id = max([r.get('id', 0) for r in self.fabrics_import_logs], default=0) + 1
			self.fabrics_import_logs.append({
				'id': log_id,
				'file_name': os.path.basename(file_path),
				'full_path': os.path.abspath(file_path),
				'records_added': records_added,
				'imported_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			})
			self.save_fabrics_import_logs()
			return log_id
		except Exception as e:
			print(f"שגיאה בהוספת לוג ייבוא: {e}")
			return None

	def delete_fabric_import_log(self, log_id: int) -> bool:
		"""מחיקת רשומת לוג לפי ID"""
		try:
			before = len(self.fabrics_import_logs)
			self.fabrics_import_logs = [r for r in self.fabrics_import_logs if r.get('id') != log_id]
			if len(self.fabrics_import_logs) != before:
				return self.save_fabrics_import_logs()
			return False
		except Exception as e:
			print(f"שגיאה במחיקת לוג ייבוא: {e}")
			return False

	def delete_fabric_import_log_and_fabrics(self, log_id: int) -> Dict[str, int]:
		"""מחיקת לוג והקבצים שייבא (רשומות מלאי) לפי import_log_id.
		:return: {'logs_deleted': int, 'fabrics_deleted': int}
		"""
		result = {'logs_deleted': 0, 'fabrics_deleted': 0}
		try:
			# מחיקת רשומות מלאי עם אותו import_log_id
			before_f = len(self.fabrics_inventory)
			self.fabrics_inventory = [r for r in self.fabrics_inventory if r.get('import_log_id') != log_id]
			after_f = len(self.fabrics_inventory)
			if after_f != before_f:
				result['fabrics_deleted'] = before_f - after_f
				self.save_fabrics_inventory()
			# מחיקת הלוג עצמו
			before_l = len(self.fabrics_import_logs)
			self.fabrics_import_logs = [r for r in self.fabrics_import_logs if r.get('id') != log_id]
			after_l = len(self.fabrics_import_logs)
			if after_l != before_l:
				result['logs_deleted'] = before_l - after_l
				self.save_fabrics_import_logs()
			return result
		except Exception as e:
			print(f"שגיאה במחיקת לוג ורשומות מלאי: {e}")
			return result

	# ===== Fabric Status Management =====
	def update_fabric_status(self, barcode: str, new_status: str) -> bool:
		"""עדכון סטטוס בד לפי ברקוד. מחזיר True אם עודכן."""
		try:
			changed = False
			for rec in self.fabrics_inventory:
				if str(rec.get('barcode')) == str(barcode):
					if rec.get('status') != new_status:
						rec['status'] = new_status
						changed = True
			if changed:
				self.save_fabrics_inventory()
			return changed
		except Exception as e:
			print(f"שגיאה בעדכון סטטוס: {e}")
			return False

	# ===== Products Catalog (New) =====
	def load_products_catalog(self) -> List[Dict]:
		"""טעינת קטלוג המוצרים המקומי."""
		try:
			if os.path.exists(self.products_catalog_file):
				with open(self.products_catalog_file, 'r', encoding='utf-8') as f:
					data = json.load(f)
					# שמירה על רשימה
					if isinstance(data, list):
						return data
			return []
		except Exception as e:
			print(f"שגיאה בטעינת קטלוג מוצרים: {e}")
			return []

	def save_products_catalog(self) -> bool:
		"""שמירת קטלוג המוצרים לקובץ."""
		try:
			with open(self.products_catalog_file, 'w', encoding='utf-8') as f:
				json.dump(self.products_catalog, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קטלוג מוצרים: {e}")
			return False

	def add_product_catalog_entry(self, name: str, size: str, fabric_type: str, fabric_color: str, print_name: str,
								 category: str = '', ticks_qty: int | str = 0, elastic_qty: int | str = 0, ribbon_qty: int | str = 0, fabric_category: str = '',
								 barcode: str = '', main_category: str = '', unit_type: str = '', zipper_qty: int | str = 0) -> int:
		"""הוספת מוצר לקטלוג עם שדות מורחבים. מחזיר ID חדש.

		:param name: שם מוצר (חובה)
		:param size: מידה
		:param fabric_type: סוג בד
		:param fabric_color: צבע בד
		:param print_name: שם פרינט
		:param category: קטגוריה (אופציונלי)
		:param ticks_qty: כמות טיקטקים (אופציונלי, מספר שלם)
		:param elastic_qty: כמות גומי (אופציונלי, מספר שלם)
		:param ribbon_qty: כמות סרט (אופציונלי, מספר שלם)
		:param zipper_qty: כמות רוכסן (אופציונלי, מספר שלם)
		"""
		try:
			if not name:
				raise ValueError("חובה להזין שם מוצר")
			def _to_int(val):
				try:
					if val in (None, ''): return 0
					return int(str(val).strip())
				except Exception:
					return 0
			ticks_i = _to_int(ticks_qty)
			elastic_i = _to_int(elastic_qty)
			ribbon_i = _to_int(ribbon_qty)
			zipper_i = _to_int(zipper_qty)
			new_id = max([p.get('id', 0) for p in self.products_catalog], default=0) + 1
			record = {
				'id': new_id,
				'name': name.strip(),
				'size': size.strip(),
				'fabric_type': fabric_type.strip(),
				'fabric_color': fabric_color.strip(),
				'print_name': print_name.strip(),
				'category': (category or '').strip(),
				'fabric_category': (fabric_category or '').strip(),
				'barcode': (barcode or '').strip(),
				'main_category': (main_category or '').strip(),
				'unit_type': (unit_type or '').strip(),
				'ticks_qty': ticks_i,
				'elastic_qty': elastic_i,
				'ribbon_qty': ribbon_i,
				'zipper_qty': zipper_i,
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.products_catalog.append(record)
			self.save_products_catalog()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת מוצר: {str(e)}")

	def delete_product_catalog_entry(self, entry_id: int) -> bool:
		"""מחיקת מוצר לפי ID. מחזיר True אם נמחק."""
		before = len(self.products_catalog)
		self.products_catalog = [p for p in self.products_catalog if int(p.get('id', 0)) != int(entry_id)]
		if len(self.products_catalog) != before:
			self.save_products_catalog(); return True
		return False

	def refresh_products_catalog(self):
		"""רענון קטלוג מהמחשב."""
		self.products_catalog = self.load_products_catalog()

	def export_products_catalog_to_excel(self, file_path: str) -> bool:
		"""ייצוא קטלוג מוצרים ל-Excel."""
		try:
			if not self.products_catalog:
				raise ValueError("אין מוצרים לייצוא")
			# בונים DataFrame מסודר עם עמודות קבועות, כולל העמודות החדשות
			columns = [
				'id','name','main_category','category','size','fabric_type','fabric_color','fabric_category',
				'print_name','barcode','unit_type','ticks_qty','elastic_qty','ribbon_qty','zipper_qty','created_at'
			]
			rows = []
			for rec in self.products_catalog:
				rows.append({
					'id': rec.get('id'),
					'name': rec.get('name',''),
					'main_category': rec.get('main_category',''),
					'category': rec.get('category',''),
					'size': rec.get('size',''),
					'fabric_type': rec.get('fabric_type',''),
					'fabric_color': rec.get('fabric_color',''),
					# ברירת מחדל: "בלי קטגוריה" אם חסר
					'fabric_category': rec.get('fabric_category') or 'בלי קטגוריה',
					'print_name': rec.get('print_name',''),
					'barcode': rec.get('barcode',''),
					'unit_type': rec.get('unit_type',''),
					'ticks_qty': rec.get('ticks_qty', 0),
					'elastic_qty': rec.get('elastic_qty', 0),
					'ribbon_qty': rec.get('ribbon_qty', 0),
					'zipper_qty': rec.get('zipper_qty', 0),
					'created_at': rec.get('created_at','')
				})
			df = pd.DataFrame(rows, columns=columns)
			df.to_excel(file_path, index=False)
			return True
		except Exception as e:
			raise Exception(f"שגיאה בייצוא קטלוג מוצרים: {str(e)}")

	def import_products_catalog_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		"""יבוא קטלוג מוצרים מ‑Excel בפורמט זהה לייצוא.

		:param file_path: נתיב לקובץ xlsx
		:param mode: 'append' להוספה לקיים (עם מניעת כפילויות), או 'overwrite' לדריסה מלאה
		:return: dict עם מונים: {'imported': N, 'skipped_duplicates': M, 'overwritten': bool}
		"""
		try:
			if not os.path.exists(file_path):
				raise Exception("קובץ לא נמצא")
			df = pd.read_excel(file_path)
			# נוודא קיום עמודות נדרשות (תמיכה לאחור: barcode/main_category/unit_type לא חובה)
			required = {'name','category','size','fabric_type','fabric_color','fabric_category','print_name','ticks_qty','elastic_qty','ribbon_qty','zipper_qty','created_at'}
			cols = {str(c).strip() for c in df.columns}
			missing = required - cols
			if missing:
				raise Exception(f"עמודות חסרות בקובץ: {', '.join(missing)}")

			# יצירת סט כפילויות קיים ע"פ מפתחות לוגיים (שם+מידה+סוג בד+צבע בד+שם פרינט)
			existing_keys = set()
			for rec in self.products_catalog:
				existing_keys.add((
					(rec.get('name') or '').strip(),
					(rec.get('size') or '').strip(),
					(rec.get('fabric_type') or '').strip(),
					(rec.get('fabric_color') or '').strip(),
					(rec.get('print_name') or '').strip()
				))

			def _to_int(x):
				try:
					if x in (None, ''):
						return 0
					return int(x)
				except Exception:
					return 0

			imported = 0
			skipped = 0
			if mode == 'overwrite':
				self.products_catalog = []
				existing_keys.clear()
				next_id = 1
				for _, row in df.iterrows():
					name = str(row.get('name') or '').strip()
					size = str(row.get('size') or '').strip()
					ft = str(row.get('fabric_type') or '').strip()
					fc = str(row.get('fabric_color') or '').strip()
					pn = str(row.get('print_name') or '').strip()
					cat = str(row.get('category') or '').strip()
					fcat = (str(row.get('fabric_category') or '').strip())
					# ננרמל 'בלי קטגוריה' לריק	
					if fcat == 'בלי קטגוריה':
						fcat = ''
					barcode = str(row.get('barcode') or '').strip() if 'barcode' in cols else ''
					main_cat = str(row.get('main_category') or '').strip() if 'main_category' in cols else ''
					unit_type = str(row.get('unit_type') or '').strip() if 'unit_type' in cols else ''
					ticks = _to_int(row.get('ticks_qty'))
					elastic = _to_int(row.get('elastic_qty'))
					ribbon = _to_int(row.get('ribbon_qty'))
					zipper = _to_int(row.get('zipper_qty'))
					created = row.get('created_at')
					created_str = str(created) if created is not None else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
					key = (name, size, ft, fc, pn)
					if key in existing_keys:
						skipped += 1; continue
					rec = {
						'id': next_id,
						'name': name,
						'size': size,
						'fabric_type': ft,
						'fabric_color': fc,
						'print_name': pn,
						'category': cat,
						'fabric_category': fcat,
						'barcode': barcode,
						'main_category': main_cat,
						'unit_type': unit_type,
						'ticks_qty': ticks,
						'elastic_qty': elastic,
						'ribbon_qty': ribbon,
						'zipper_qty': zipper,
						'created_at': created_str
					}
					self.products_catalog.append(rec)
					existing_keys.add(key)
					next_id += 1; imported += 1
			else:
				# append
				next_id = max([p.get('id',0) for p in self.products_catalog], default=0) + 1
				for _, row in df.iterrows():
					name = str(row.get('name') or '').strip()
					size = str(row.get('size') or '').strip()
					ft = str(row.get('fabric_type') or '').strip()
					fc = str(row.get('fabric_color') or '').strip()
					pn = str(row.get('print_name') or '').strip()
					cat = str(row.get('category') or '').strip()
					fcat = (str(row.get('fabric_category') or '').strip())
					if fcat == 'בלי קטגוריה':
						fcat = ''
					barcode = str(row.get('barcode') or '').strip() if 'barcode' in cols else ''
					main_cat = str(row.get('main_category') or '').strip() if 'main_category' in cols else ''
					unit_type = str(row.get('unit_type') or '').strip() if 'unit_type' in cols else ''
					ticks = _to_int(row.get('ticks_qty'))
					elastic = _to_int(row.get('elastic_qty'))
					ribbon = _to_int(row.get('ribbon_qty'))
					zipper = _to_int(row.get('zipper_qty'))
					created = row.get('created_at')
					created_str = str(created) if created is not None else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
					key = (name, size, ft, fc, pn)
					if key in existing_keys:
						skipped += 1; continue
					rec = {
						'id': next_id,
						'name': name,
						'size': size,
						'fabric_type': ft,
						'fabric_color': fc,
						'print_name': pn,
						'category': cat,
						'fabric_category': fcat,
						'barcode': barcode,
						'main_category': main_cat,
						'unit_type': unit_type,
						'ticks_qty': ticks,
						'elastic_qty': elastic,
						'ribbon_qty': ribbon,
						'zipper_qty': zipper,
						'created_at': created_str
					}
					self.products_catalog.append(rec)
					existing_keys.add(key)
					next_id += 1; imported += 1

			self.save_products_catalog()
			return {'imported': imported, 'skipped_duplicates': skipped, 'overwritten': mode == 'overwrite'}
		except Exception as e:
			raise Exception(f"שגיאה בייבוא קטלוג מוצרים: {str(e)}")

	# ===== Sewing Accessories =====
	def load_sewing_accessories(self) -> List[Dict]:
		"""טעינת רשימת אביזרי תפירה."""
		try:
			if os.path.exists(self.sewing_accessories_file):
				with open(self.sewing_accessories_file, 'r', encoding='utf-8') as f:
					data = json.load(f)
					if isinstance(data, list):
						return data
			return []
		except Exception as e:
			print(f"שגיאה בטעינת אביזרי תפירה: {e}"); return []

	def save_sewing_accessories(self) -> bool:
		try:
			with open(self.sewing_accessories_file, 'w', encoding='utf-8') as f:
				json.dump(self.sewing_accessories, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת אביזרי תפירה: {e}"); return False

	def add_sewing_accessory(self, name: str, unit: str) -> int:
		"""הוספת אביזר תפירה (שם + יחידת מדידה)."""
		try:
			if not name:
				raise ValueError("חובה להזין שם אביזר")
			# מניעת כפילות שם + יחידה
			for rec in self.sewing_accessories:
				if rec.get('name','').strip() == name.strip() and rec.get('unit','').strip() == (unit or '').strip():
					raise ValueError("האביזר כבר קיים")
			new_id = max([r.get('id',0) for r in self.sewing_accessories], default=0) + 1
			rec = {
				'id': new_id,
				'name': name.strip(),
				'unit': (unit or '').strip(),
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.sewing_accessories.append(rec)
			self.save_sewing_accessories(); return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת אביזר: {str(e)}")

	def delete_sewing_accessory(self, acc_id: int) -> bool:
		before = len(self.sewing_accessories)
		self.sewing_accessories = [r for r in self.sewing_accessories if int(r.get('id',0)) != int(acc_id)]
		if len(self.sewing_accessories) != before:
			self.save_sewing_accessories(); return True
		return False

	def refresh_sewing_accessories(self):
		self.sewing_accessories = self.load_sewing_accessories()

	# ===== Categories Management =====
	def load_categories(self) -> List[Dict]:
		try:
			if os.path.exists(self.categories_file):
				with open(self.categories_file, 'r', encoding='utf-8') as f:
					data = json.load(f)
					if isinstance(data, list):
						return data
			return []
		except Exception as e:
			print(f"שגיאה בטעינת קטגוריות: {e}"); return []

	def save_categories(self) -> bool:
		try:
			with open(self.categories_file, 'w', encoding='utf-8') as f:
				json.dump(self.categories, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קטגוריות: {e}"); return False

	# ===== Main Categories Management =====
	def load_main_categories(self) -> List[Dict]:
		try:
			if os.path.exists(self.main_categories_file):
				with open(self.main_categories_file, 'r', encoding='utf-8') as f:
					data = json.load(f)
					if isinstance(data, list):
						return data
			return []
		except Exception as e:
			print(f"שגיאה בטעינת קטגוריות ראשיות: {e}"); return []

	def save_main_categories(self) -> bool:
		try:
			with open(self.main_categories_file, 'w', encoding='utf-8') as f:
				json.dump(self.main_categories, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קטגוריות ראשיות: {e}"); return False

	def add_main_category(self, name: str) -> int:
		try:
			if not name:
				raise ValueError("חובה להזין שם קטגוריה ראשית")
			for c in self.main_categories:
				if c.get('name','').strip() == name.strip():
					raise ValueError("קטגוריה ראשית קיימת")
			new_id = max([c.get('id',0) for c in self.main_categories], default=0) + 1
			rec = {
				'id': new_id,
				'name': name.strip(),
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
				'fields': []
			}
			self.main_categories.append(rec)
			self.save_main_categories(); return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קטגוריה ראשית: {str(e)}")

	def delete_main_category(self, cat_id: int) -> bool:
		before = len(self.main_categories)
		self.main_categories = [c for c in self.main_categories if int(c.get('id',0)) != int(cat_id)]
		if len(self.main_categories) != before:
			self.save_main_categories(); return True
		return False

	def refresh_main_categories(self):
		self.main_categories = self.load_main_categories()

	def get_main_category_fields(self, cat_id: int) -> list[str]:
		try:
			for c in self.main_categories:
				if int(c.get('id', 0)) == int(cat_id):
					fields = c.get('fields')
					return list(fields) if isinstance(fields, list) else []
			return []
		except Exception:
			return []

	def set_main_category_fields(self, cat_id: int, fields: list[str]) -> bool:
		try:
			for c in self.main_categories:
				if int(c.get('id', 0)) == int(cat_id):
					c['fields'] = list(fields)
					self.save_main_categories()
					return True
		except Exception:
			pass
		return False

	def add_category(self, name: str) -> int:
		try:
			if not name:
				raise ValueError("חובה להזין שם קטגוריה")
			for c in self.categories:
				if c.get('name','').strip() == name.strip():
					raise ValueError("קטגוריה קיימת")
			new_id = max([c.get('id',0) for c in self.categories], default=0) + 1
			rec = {'id': new_id, 'name': name.strip(), 'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
			self.categories.append(rec)
			self.save_categories(); return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קטגוריה: {str(e)}")

	def delete_category(self, cat_id: int) -> bool:
		before = len(self.categories)
		self.categories = [c for c in self.categories if int(c.get('id',0)) != int(cat_id)]
		if len(self.categories) != before:
			self.save_categories(); return True
		return False

	def refresh_categories(self):
		self.categories = self.load_categories()

	# ===== Product Attribute Lists (sizes, fabric types, colors, print names) =====
	def _load_simple_list(self, path: str) -> list[dict]:
		try:
			if os.path.exists(path):
				with open(path, 'r', encoding='utf-8') as f:
					data = json.load(f)
					if isinstance(data, list):
						return data
			return []
		except Exception as e:
			print(f"שגיאה בטעינת קובץ {path}: {e}"); return []

	def _save_simple_list(self, path: str, data: list[dict]) -> bool:
		try:
			with open(path, 'w', encoding='utf-8') as f:
				json.dump(data, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קובץ {path}: {e}"); return False

	def load_product_sizes(self):
		return self._load_simple_list(self.product_sizes_file)

	def load_fabric_types(self):
		return self._load_simple_list(self.fabric_types_file)

	def load_fabric_colors(self):
		return self._load_simple_list(self.fabric_colors_file)

	def load_print_names(self):
		return self._load_simple_list(self.print_names_file)

	def load_fabric_categories(self):
		return self._load_simple_list(self.fabric_categories_file)

	def load_model_names(self):
		return self._load_simple_list(self.model_names_file)

	def save_product_sizes(self):
		return self._save_simple_list(self.product_sizes_file, self.product_sizes)

	def save_fabric_types(self):
		return self._save_simple_list(self.fabric_types_file, self.product_fabric_types)

	def save_fabric_colors(self):
		return self._save_simple_list(self.fabric_colors_file, self.product_fabric_colors)

	def save_print_names(self):
		return self._save_simple_list(self.print_names_file, self.product_print_names)

	def save_fabric_categories(self):
		return self._save_simple_list(self.fabric_categories_file, self.product_fabric_categories)

	def save_model_names(self):
		return self._save_simple_list(self.model_names_file, self.product_model_names)

	def _add_to_simple_list(self, data_list: list[dict], save_func, name: str, extra: dict | None = None) -> int:
		if not name:
			raise Exception("חובה להזין שם")
		for rec in data_list:
			if rec.get('name','').strip() == name.strip():
				raise Exception("פריט כבר קיים")
		new_id = max([r.get('id',0) for r in data_list], default=0) + 1
		rec = {'id': new_id, 'name': name.strip(), 'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
		try:
			if extra and isinstance(extra, dict):
				rec.update(extra)
		except Exception:
			pass
		data_list.append(rec)
		save_func(); return new_id

	def _delete_from_simple_list(self, data_list: list[dict], save_func, rec_id: int) -> bool:
		before = len(data_list)
		data_list[:] = [r for r in data_list if int(r.get('id',0)) != int(rec_id)]
		if len(data_list) != before:
			save_func(); return True
		return False

	def add_product_size(self, name: str, main_category: str = '') -> int:
		return self._add_to_simple_list(self.product_sizes, self.save_product_sizes, name, {'main_category': (main_category or '').strip()})

	def delete_product_size(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_sizes, self.save_product_sizes, rec_id)

	def add_fabric_type_item(self, name: str, main_category: str | list[str] = '') -> int:
		"""Add fabric type; supports single or multiple main categories.
		- main_category: str or list[str]. If list, will be saved under 'main_categories' and 'main_category' (first) for compatibility.
		"""
		extra: dict = {}
		mcs: list[str] = []
		try:
			if isinstance(main_category, list):
				mcs = [ (m or '').strip() for m in main_category if (m or '').strip() ]
			elif isinstance(main_category, str):
				s = (main_category or '').strip()
				# support comma-separated input
				mcs = [t.strip() for t in s.split(',') if t.strip()] if s else []
			else:
				mcs = []
		except Exception:
			mcs = []
		if mcs:
			extra['main_categories'] = mcs
			extra['main_category'] = mcs[0]
		else:
			extra['main_category'] = (main_category or '').strip() if isinstance(main_category, str) else ''
		return self._add_to_simple_list(self.product_fabric_types, self.save_fabric_types, name, extra)

	def delete_fabric_type_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_types, self.save_fabric_types, rec_id)

	def add_fabric_color_item(self, name: str, main_category: str | list[str] = '') -> int:
		"""Add fabric color; supports single or multiple main categories."""
		extra: dict = {}
		mcs: list[str] = []
		try:
			if isinstance(main_category, list):
				mcs = [ (m or '').strip() for m in main_category if (m or '').strip() ]
			elif isinstance(main_category, str):
				s = (main_category or '').strip()
				mcs = [t.strip() for t in s.split(',') if t.strip()] if s else []
			else:
				mcs = []
		except Exception:
			mcs = []
		if mcs:
			extra['main_categories'] = mcs
			extra['main_category'] = mcs[0]
		else:
			extra['main_category'] = (main_category or '').strip() if isinstance(main_category, str) else ''
		return self._add_to_simple_list(self.product_fabric_colors, self.save_fabric_colors, name, extra)

	def delete_fabric_color_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_colors, self.save_fabric_colors, rec_id)

	def add_print_name_item(self, name: str, main_category: str | list[str] = '') -> int:
		"""Add print name; supports single or multiple main categories."""
		extra: dict = {}
		mcs: list[str] = []
		try:
			if isinstance(main_category, list):
				mcs = [ (m or '').strip() for m in main_category if (m or '').strip() ]
			elif isinstance(main_category, str):
				s = (main_category or '').strip()
				mcs = [t.strip() for t in s.split(',') if t.strip()] if s else []
			else:
				mcs = []
		except Exception:
			mcs = []
		if mcs:
			extra['main_categories'] = mcs
			extra['main_category'] = mcs[0]
		else:
			extra['main_category'] = (main_category or '').strip() if isinstance(main_category, str) else ''
		return self._add_to_simple_list(self.product_print_names, self.save_print_names, name, extra)

	def delete_print_name_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_print_names, self.save_print_names, rec_id)

	def add_fabric_category_item(self, name: str, main_category: str = '') -> int:
		return self._add_to_simple_list(self.product_fabric_categories, self.save_fabric_categories, name, {'main_category': (main_category or '').strip()})

	def delete_fabric_category_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_categories, self.save_fabric_categories, rec_id)

	def add_model_name_item(self, name: str, main_category: str = '') -> int:
		return self._add_to_simple_list(self.product_model_names, self.save_model_names, name, {'main_category': (main_category or '').strip()})

	def delete_model_name_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_model_names, self.save_model_names, rec_id)

	def refresh_product_attributes(self):
		self.product_sizes = self.load_product_sizes()
		self.product_fabric_types = self.load_fabric_types()
		self.product_fabric_colors = self.load_fabric_colors()
		self.product_print_names = self.load_print_names()
		self.product_fabric_categories = self.load_fabric_categories()
		self.product_model_names = self.load_model_names()
		try:
			self._ensure_main_category_on_attributes()
		except Exception:
			pass

	def _ensure_main_category_on_attributes(self, default: str = 'בגדים'):
		"""Ensure each attribute record has 'main_category' and optional 'main_categories'.
		- If neither present, default to provided value.
		- If only 'main_category' present, mirror into 'main_categories' list.
		- If 'main_category' contains a comma-separated list (legacy data), split into list and
		  set the first token as the primary.
		- If 'main_categories' exists but contains a single string with commas (bad import),
		  split that string into a proper list and set the first as primary.
		"""
		changed = False
		for lst, saver in [
			(self.product_sizes, self.save_product_sizes),
			(self.product_fabric_types, self.save_fabric_types),
			(self.product_fabric_colors, self.save_fabric_colors),
			(self.product_print_names, self.save_print_names),
			(self.product_fabric_categories, self.save_fabric_categories),
			(self.product_model_names, self.save_model_names),
		]:
			try:
				for rec in lst:
					mc = (rec.get('main_category') or '').strip()
					mcs = rec.get('main_categories') if isinstance(rec.get('main_categories'), list) else []
					# If main_categories exists but is a single comma-joined string in a list, fix it
					if mcs and isinstance(mcs, list) and len(mcs) == 1 and isinstance(mcs[0], str) and (',' in mcs[0]):
						parts = [t.strip() for t in mcs[0].split(',') if t and t.strip()]
						if parts:
							rec['main_categories'] = parts
							rec['main_category'] = parts[0]
							changed = True
					if not mc and not mcs:
						rec['main_category'] = default
						rec['main_categories'] = [default]
						changed = True
					elif mc and not mcs:
						# Normalize legacy comma-separated primary into list
						parts = [t.strip() for t in mc.split(',') if t and t.strip()]
						if parts:
							rec['main_categories'] = parts
							rec['main_category'] = parts[0]
						else:
							rec['main_categories'] = [mc]
						changed = True
			except Exception:
				pass
		if changed:
			# Save each list via its saver to persist defaults
			try: self.save_product_sizes()
			except Exception: pass
			try: self.save_fabric_types()
			except Exception: pass
			try: self.save_fabric_colors()
			except Exception: pass
			try: self.save_print_names()
			except Exception: pass
			try: self.save_fabric_categories()
			except Exception: pass
			try: self.save_model_names()
			except Exception: pass

	# ===== Attribute export/import (Excel) =====
	def _export_attr_list(self, items: list[dict], file_path: str) -> bool:
		"""Export attribute list to Excel with one category column: id,name,main_category,created_at.
		- If multiple categories exist, they are joined by commas into main_category for user editing.
		"""
		try:
			rows = []
			for rec in items or []:
				mcs = rec.get('main_categories') if isinstance(rec.get('main_categories'), list) else None
				mc = (rec.get('main_category') or '').strip()
				if not mcs and mc:
					mcs = [mc]
				rows.append({
					'id': rec.get('id'),
					'name': rec.get('name', ''),
					'main_category': ','.join(mcs) if mcs else '',
					'created_at': rec.get('created_at',''),
				})
			df = pd.DataFrame(rows, columns=['id','name','main_category','created_at'])
			df.to_excel(file_path, index=False)
			return True
		except Exception as e:
			raise Exception(f"שגיאה בייצוא אקסל: {e}")

	def _import_attr_list(self, file_path: str, mode: str, target_list: list[dict], save_func, add_func, supports_multi_mc: bool = False) -> dict:
		"""Import attribute list from Excel. If overwrite, replace list; else append unique by 'name'.
		- supports_multi_mc: when True, will pass list of categories to add_func; otherwise single string.
		"""
		try:
			if not os.path.exists(file_path):
				raise Exception("קובץ לא נמצא")
			df = pd.read_excel(file_path)
			cols = {str(c).strip() for c in df.columns}
			if 'name' not in cols:
				raise Exception("עמודת 'name' חסרה בקובץ")
			imported = 0; skipped = 0
			if mode == 'overwrite':
				target_list.clear()
				save_func()
				existing = set()
			else:
				existing = { (rec.get('name') or '').strip() for rec in (target_list or []) }
			for _, row in df.iterrows():
				name = str(row.get('name') or '').strip()
				if not name:
					continue
				if name in existing:
					skipped += 1; continue
				# categories
				mcs = []
				try:
					if 'main_categories' in cols:
						val = row.get('main_categories')
						s = '' if val is None else str(val)
						mcs = [t.strip() for t in s.split(',') if t and t.strip()]
					if (not mcs) and 'main_category' in cols:
						mc_val = row.get('main_category')
						s = '' if mc_val is None else str(mc_val)
						# When multi-category is supported, split comma-separated values from single column
						if supports_multi_mc:
							mcs = [t.strip() for t in s.split(',') if t and t.strip()]
						else:
							mc = s.strip()
							if mc: mcs = [mc]
				except Exception:
					mcs = []
				try:
					if supports_multi_mc:
						add_func(name, main_category=(mcs or ''))
					else:
						mc = mcs[0] if mcs else ''
						add_func(name, main_category=mc)
					imported += 1
					existing.add(name)
				except Exception:
					skipped += 1
			save_func(); return {'imported': imported, 'skipped_duplicates': skipped, 'overwritten': (mode=='overwrite')}
		except Exception as e:
			raise Exception(f"שגיאה בייבוא אקסל: {e}")

	# Sizes
	def export_sizes_to_excel(self, file_path: str) -> bool:
		return self._export_attr_list(self.product_sizes, file_path)

	def import_sizes_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		return self._import_attr_list(file_path, mode, self.product_sizes, self.save_product_sizes, self.add_product_size, supports_multi_mc=False)

	# Fabric Types
	def export_fabric_types_to_excel(self, file_path: str) -> bool:
		return self._export_attr_list(self.product_fabric_types, file_path)

	def import_fabric_types_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		return self._import_attr_list(file_path, mode, self.product_fabric_types, self.save_fabric_types, self.add_fabric_type_item, supports_multi_mc=True)

	# Fabric Colors
	def export_fabric_colors_to_excel(self, file_path: str) -> bool:
		return self._export_attr_list(self.product_fabric_colors, file_path)

	def import_fabric_colors_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		return self._import_attr_list(file_path, mode, self.product_fabric_colors, self.save_fabric_colors, self.add_fabric_color_item, supports_multi_mc=True)

	# Print Names
	def export_print_names_to_excel(self, file_path: str) -> bool:
		return self._export_attr_list(self.product_print_names, file_path)

	def import_print_names_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		return self._import_attr_list(file_path, mode, self.product_print_names, self.save_print_names, self.add_print_name_item, supports_multi_mc=True)

	# Model Names
	def export_model_names_to_excel(self, file_path: str) -> bool:
		return self._export_attr_list(self.product_model_names, file_path)

	def import_model_names_from_excel(self, file_path: str, mode: str = 'append') -> dict:
		return self._import_attr_list(file_path, mode, self.product_model_names, self.save_model_names, self.add_model_name_item, supports_multi_mc=False)

