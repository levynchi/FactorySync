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
    
	def __init__(self, drawings_file: str = "drawings_data.json", returned_drawings_file: str = "returned_drawings.json", fabrics_inventory_file: str = "fabrics_inventory.json", fabrics_imports_file: str = "fabrics_import_logs.json", supplier_receipts_file: str = "supplier_receipts.json", products_catalog_file: str = "products_catalog.json"):
		self.drawings_file = drawings_file
		# קובץ לקליטת ציורים שחזרו מייצור
		self.returned_drawings_file = returned_drawings_file
		# קובץ מלאי בדים
		self.fabrics_inventory_file = fabrics_inventory_file
		# קובץ לוג של ייבוא קבצי מלאי בדים
		self.fabrics_imports_file = fabrics_imports_file
		# קובץ קליטות מספק (הזנה ידנית של מוצרים וכמויות)
		self.supplier_receipts_file = supplier_receipts_file
		# קובץ קטלוג מוצרים (חדש)
		self.products_catalog_file = products_catalog_file
		self.drawings_data = self.load_drawings_data()
		self.returned_drawings_data = self.load_returned_drawings_data()
		self.fabrics_inventory = self.load_fabrics_inventory()
		self.fabrics_import_logs = self.load_fabrics_import_logs()
		self.supplier_receipts = self.load_supplier_receipts()
		self.products_catalog = self.load_products_catalog()
    
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
    
	# ===== Returned Drawings Handling =====
	def load_returned_drawings_data(self) -> List[Dict]:
		"""טעינת נתוני ציורים שחזרו מייצור"""
		try:
			if os.path.exists(self.returned_drawings_file):
				with open(self.returned_drawings_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת ציורים חוזרים: {e}")
			return []
    
	def save_returned_drawings_data(self) -> bool:
		"""שמירת נתוני ציורים שחזרו"""
		try:
			with open(self.returned_drawings_file, 'w', encoding='utf-8') as f:
				json.dump(self.returned_drawings_data, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת ציורים חוזרים: {e}")
			return False

	# ===== Supplier Receipts (Manual Products Intake) =====
	def load_supplier_receipts(self) -> List[Dict]:
		"""טעינת קליטות מספק"""
		try:
			if os.path.exists(self.supplier_receipts_file):
				with open(self.supplier_receipts_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת קליטות ספק: {e}")
			return []

	def save_supplier_receipts(self) -> bool:
		"""שמירת קליטות מספק"""
		try:
			with open(self.supplier_receipts_file, 'w', encoding='utf-8') as f:
				json.dump(self.supplier_receipts, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת קליטות ספק: {e}")
			return False

	def add_supplier_receipt(self, supplier: str, date_str: str, lines: List[Dict]) -> int:
		"""הוספת קליטה מספק.
		:param supplier: שם ספק
		:param date_str: תאריך (YYYY-MM-DD)
		:param lines: רשימת שורות: {product, size, quantity, note}
		"""
		try:
			if not supplier:
				raise ValueError("חסר שם ספק")
			if not lines:
				raise ValueError("אין שורות לקליטה")
			new_id = max([r.get('id', 0) for r in self.supplier_receipts], default=0) + 1
			total_quantity = sum(int(l.get('quantity', 0)) for l in lines)
			receipt = {
				'id': new_id,
				'supplier': supplier,
				'date': date_str,
				'lines': lines,
				'total_quantity': total_quantity,
				'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
			}
			self.supplier_receipts.append(receipt)
			self.save_supplier_receipts()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קליטת ספק: {str(e)}")

	def refresh_supplier_receipts(self):
		self.supplier_receipts = self.load_supplier_receipts()
    
	def add_returned_drawing(self, drawing_id: str, date_str: str, barcodes: List[str], source: str = None, layers: int = None) -> int:
		"""הוספת קליטת ציור חוזר
		:param drawing_id: מזהה הציור (טקסט / מספר)
		:param date_str: תאריך הקליטה (YYYY-MM-DD)
		:param barcodes: רשימת ברקודים שנקלטו
		:return: ID פנימי של הרשומה החדשה
		"""
		try:
			if not barcodes:
				raise ValueError("לא נקלטו ברקודים")
			# יצירת מזהה חדש
			new_id = max([r.get('id', 0) for r in self.returned_drawings_data], default=0) + 1
			record = {
				'id': new_id,
				'drawing_id': drawing_id,
				'date': date_str,
				'barcodes': barcodes,
				'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
			}
			if source:
				record['source'] = source
			if layers is not None:
				record['layers'] = layers
			self.returned_drawings_data.append(record)
			self.save_returned_drawings_data()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת ציור חוזר: {str(e)}")
    
	def get_returned_drawings_summary(self) -> Dict[str, Any]:
		"""סיכום מהיר של הקליטות"""
		total = len(self.returned_drawings_data)
		total_barcodes = sum(len(r.get('barcodes', [])) for r in self.returned_drawings_data)
		return {
			'total_records': total,
			'total_barcodes': total_barcodes
		}
    
	def refresh_returned_drawings(self):
		"""רענון נתוני ציורים חוזרים"""
		self.returned_drawings_data = self.load_returned_drawings_data()

	def delete_returned_drawing(self, record_id: int) -> bool:
		"""מחיקת רשומת ציור חוזר לפי ID.
		:return: True אם נמחקה רשומה.
		"""
		before = len(self.returned_drawings_data)
		self.returned_drawings_data = [r for r in self.returned_drawings_data if int(r.get('id', 0)) != int(record_id)]
		if len(self.returned_drawings_data) != before:
			self.save_returned_drawings_data()
			return True
		return False
    
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
    
	def add_to_local_table(self, results: List[Dict], file_name: str = "", fabric_type: str = "") -> int:
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
			self.drawings_data = [r for r in self.drawings_data if r.get('id') != drawing_id]
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

	def add_product_catalog_entry(self, name: str, size: str, fabric_type: str, fabric_color: str, print_name: str) -> int:
		"""הוספת מוצר לקטלוג. מחזיר ID חדש."""
		try:
			if not name:
				raise ValueError("חובה להזין שם מוצר")
			new_id = max([p.get('id', 0) for p in self.products_catalog], default=0) + 1
			record = {
				'id': new_id,
				'name': name.strip(),
				'size': size.strip(),
				'fabric_type': fabric_type.strip(),
				'fabric_color': fabric_color.strip(),
				'print_name': print_name.strip(),
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
			df = pd.DataFrame(self.products_catalog)
			df.to_excel(file_path, index=False)
			return True
		except Exception as e:
			raise Exception(f"שגיאה בייצוא קטלוג מוצרים: {str(e)}")

