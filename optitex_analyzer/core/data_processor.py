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
    
	def __init__(self, drawings_file: str = "drawings_data.json", fabrics_inventory_file: str = "fabrics_inventory.json", fabrics_imports_file: str = "fabrics_import_logs.json", supplier_receipts_file: str = "supplier_receipts.json", products_catalog_file: str = "products_catalog.json", suppliers_file: str = "suppliers.json", supplier_intakes_file: str = "supplier_intakes.json", delivery_notes_file: str = "delivery_notes.json", sewing_accessories_file: str = "sewing_accessories.json", categories_file: str = "categories.json", product_sizes_file: str = "product_sizes.json", fabric_types_file: str = "fabric_types.json", fabric_colors_file: str = "fabric_colors.json", print_names_file: str = "print_names.json", fabric_categories_file: str = "fabric_categories.json", model_names_file: str = "model_names.json", main_categories_file: str = "main_categories.json"):
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
		self.suppliers = self.load_suppliers()
		# Fabrics shipments (שליחת בדים)
		self.fabrics_shipments_file = 'fabrics_shipments.json'
		self.fabrics_shipments = self._load_json_list(self.fabrics_shipments_file)
		# Fabrics intakes (קליטת בדים) + Unbarcoded fabrics
		self.fabrics_intakes_file = 'fabrics_intakes.json'
		self.fabrics_intakes = self._load_json_list(self.fabrics_intakes_file)
		self.fabrics_unbarcoded_file = 'fabrics_unbarcoded.json'
		self.fabrics_unbarcoded = self._load_json_list(self.fabrics_unbarcoded_file)
		# load split receipts (may be empty on first run)
		self.supplier_intakes = self._load_json_list(self.supplier_intakes_file)
		self.delivery_notes = self._load_json_list(self.delivery_notes_file)
		# migration from legacy combined file if needed
		self._migrate_legacy_supplier_receipts()
		# combined view (backward compatibility for old UI code)
		self.supplier_receipts = self.supplier_intakes + self.delivery_notes
		# Barcodes data
		self.barcodes_data_file = 'barcodes_data.json'
		self.barcodes_data = self.load_barcodes_data()

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

	def load_barcodes_data(self) -> Dict:
		"""טעינת נתוני ברקודים"""
		try:
			if os.path.exists(self.barcodes_data_file):
				with open(self.barcodes_data_file, 'r', encoding='utf-8') as f:
					return json.load(f)
			# Default barcode data
			return {
				'last_barcode': '7297555019592',
				'last_updated': datetime.now().isoformat()
			}
		except Exception as e:
			print(f"שגיאה בטעינת נתוני ברקודים: {e}")
			return {
				'last_barcode': '7297555019592',
				'last_updated': datetime.now().isoformat()
			}

	def save_barcodes_data(self, last_barcode: str) -> bool:
		"""שמירת נתוני ברקודים"""
		try:
			data = {
				'last_barcode': last_barcode,
				'last_updated': datetime.now().isoformat()
			}
			with open(self.barcodes_data_file, 'w', encoding='utf-8') as f:
				json.dump(data, f, indent=2, ensure_ascii=False)
			self.barcodes_data = data
			return True
		except Exception as e:
			print(f"שגיאה בשמירת נתוני ברקודים: {e}")
			return False

	def calculate_ean13_checksum(self, base_12_digits: str) -> str:
		"""
		Calculate EAN-13 checksum digit.
		
		Args:
			base_12_digits: String of 12 digits
			
		Returns:
			Single digit checksum (0-9)
		"""
		if len(base_12_digits) != 12:
			raise ValueError("Base must be exactly 12 digits")
		
		# Sum odd positions (1st, 3rd, 5th, etc. - index 0, 2, 4...)
		odd_sum = sum(int(base_12_digits[i]) for i in range(0, 12, 2))
		
		# Sum even positions (2nd, 4th, 6th, etc. - index 1, 3, 5...) and multiply by 3
		even_sum = sum(int(base_12_digits[i]) for i in range(1, 12, 2)) * 3
		
		# Total sum
		total = odd_sum + even_sum
		
		# Checksum is (10 - (total mod 10)) mod 10
		checksum = (10 - (total % 10)) % 10
		
		return str(checksum)

	def generate_next_barcode(self, current_barcode: str) -> str:
		"""
		Generate the next barcode in sequence with proper EAN-13 checksum.
		
		Args:
			current_barcode: Current 13-digit EAN-13 barcode
			
		Returns:
			Next 13-digit EAN-13 barcode
		"""
		if len(current_barcode) != 13:
			raise ValueError("Barcode must be exactly 13 digits")
		
		# Remove checksum digit (last digit)
		base_12 = current_barcode[:12]
		
		# Increment the base
		base_number = int(base_12) + 1
		
		# Pad back to 12 digits
		new_base_12 = str(base_number).zfill(12)
		
		# Calculate new checksum
		checksum = self.calculate_ean13_checksum(new_base_12)
		
		# Return full 13-digit barcode
		return new_base_12 + checksum
    
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

	def add_supplier_receipt(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, accessories: List[Dict] | None = None, receipt_kind: str = "supplier_intake") -> int:
		"""שכבת תאימות – מפנה לפונקציה המתאימה לפי receipt_kind."""
		if receipt_kind == 'delivery_note':
			return self.add_delivery_note(supplier, date_str, lines, packages, accessories)
		return self.add_supplier_intake(supplier, date_str, lines, packages)

	def _next_id(self, records: List[Dict]) -> int:
		return max([r.get('id', 0) for r in records], default=0) + 1

	def add_supplier_intake(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, arrival_date: str = '', supplier_doc_number: str = '') -> int:
		try:
			if not supplier: raise ValueError("חסר שם ספק")
			if not lines: raise ValueError("אין שורות לקליטה")
			new_id = self._next_id(self.supplier_intakes)
			total_quantity = sum(int(l.get('quantity',0)) for l in lines)
			record = {
				'id': new_id,
				'supplier': supplier,
				'date': date_str,
				'arrival_date': arrival_date,
				'supplier_doc_number': supplier_doc_number,
				'lines': lines,
				'total_quantity': total_quantity,
				'packages': packages or [],
				'receipt_kind': 'supplier_intake',
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.supplier_intakes.append(record)
			self._save_json_list(self.supplier_intakes_file, self.supplier_intakes)
			self._rebuild_combined_receipts()
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קליטת ספק: {e}")

	def add_delivery_note(self, supplier: str, date_str: str, lines: List[Dict], packages: List[Dict] | None = None, accessories: List[Dict] | None = None) -> int:
		try:
			if not supplier: raise ValueError("חסר שם ספק")
			if not lines and not accessories: raise ValueError("אין שורות לקליטה או אביזרי תפירה")
			new_id = self._next_id(self.delivery_notes)
			total_quantity = sum(int(l.get('quantity',0)) for l in lines)
			record = {
				'id': new_id,
				'supplier': supplier,
				'date': date_str,
				'lines': lines,
				'total_quantity': total_quantity,
				'packages': packages or [],
				'accessories': accessories or [],
				'receipt_kind': 'delivery_note',
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

	# ===== Fabrics Shipments (שליחת בדים) =====
	def refresh_fabrics_shipments(self):
		"""Reload fabrics shipments from disk."""
		try:
			self.fabrics_shipments = self._load_json_list(self.fabrics_shipments_file)
		except Exception:
			self.fabrics_shipments = []

	def add_fabrics_shipment(self, barcodes: list[str], packages: list[dict] | None = None, date_str: str = '', fabric_type: str = '', color_name: str = '', color_no: str = '', net_kg: float | int | str = 0, meters: float | int | str = 0, supplier: str = '') -> int:
		"""Create a fabrics shipment document and return its new ID.

		barcodes: list of barcode strings included in this shipment
		packages: list of {'package_type','quantity','driver'} dicts
		date_str: optional date string; defaults to today if empty
		fabric_type/color_name/color_no/net_kg/meters: optional summary fields
		supplier: optional supplier name the fabrics are being sent to
		"""
		try:
			if not barcodes:
				raise ValueError("אין ברקודים לשמירה")
			new_id = max([r.get('id', 0) for r in self.fabrics_shipments], default=0) + 1
			def _to_float(x):
				try:
					if x in (None, ''): return 0.0
					return float(str(x).replace(',', '.'))
				except Exception:
					return 0.0
			rec = {
				'id': new_id,
				'date': date_str or datetime.now().strftime('%Y-%m-%d'),
				'barcodes': list(barcodes),
				'count_barcodes': len(barcodes),
				'packages': packages or [],
				'fabric_type': (fabric_type or ''),
				'color_name': (color_name or ''),
				'color_no': (color_no or ''),
				'net_kg': _to_float(net_kg),
				'meters': _to_float(meters),
				'supplier': supplier or '',
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.fabrics_shipments.append(rec)
			self._save_json_list(self.fabrics_shipments_file, self.fabrics_shipments)
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת שליחת בדים: {e}")

	def delete_fabrics_shipment(self, shipment_id: int) -> bool:
		"""Delete a fabrics shipment by ID."""
		before = len(self.fabrics_shipments)
		try:
			self.fabrics_shipments = [r for r in self.fabrics_shipments if int(r.get('id', -1)) != int(shipment_id)]
		except Exception:
			self.fabrics_shipments = [r for r in self.fabrics_shipments if (r.get('id') != shipment_id)]
		if len(self.fabrics_shipments) != before:
			self._save_json_list(self.fabrics_shipments_file, self.fabrics_shipments)
			return True
		return False

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

	# ===== Fabrics Intakes (קליטת בדים) =====
	def refresh_fabrics_intakes(self):
		"""Reload fabrics intakes list from disk."""
		self.fabrics_intakes = self._load_json_list(self.fabrics_intakes_file)

	def add_fabrics_intake(self, barcodes: list[str] | None = None, packages: list[dict] | None = None, supplier: str = '', date_str: str = '', unbarcoded_items: list[dict] | None = None) -> int:
		"""Save a Fabrics Intake document and return its new ID.

		Fields stored match fabrics_intakes.json schema used by the UI list tab.
		"""
		try:
			barcodes = list(barcodes or [])
			packages = list(packages or [])
			unbarcoded_items = list(unbarcoded_items or [])
			new_id = max([r.get('id', 0) for r in self.fabrics_intakes], default=0) + 1
			rec = {
				'id': new_id,
				'date': date_str or datetime.now().strftime('%Y-%m-%d'),
				'supplier': supplier or '',
				'barcodes': barcodes,
				'count_barcodes': len(barcodes),
				'unbarcoded_items': unbarcoded_items if unbarcoded_items else [],
				'packages': packages,
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.fabrics_intakes.append(rec)
			self._save_json_list(self.fabrics_intakes_file, self.fabrics_intakes)
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת קליטת בדים: {e}")

	def delete_fabrics_intake(self, rec_id: int) -> bool:
		"""Delete a fabrics intake by ID."""
		before = len(self.fabrics_intakes)
		try:
			self.fabrics_intakes = [r for r in self.fabrics_intakes if int(r.get('id', -1)) != int(rec_id)]
		except Exception:
			self.fabrics_intakes = [r for r in self.fabrics_intakes if (r.get('id') != rec_id)]
		if len(self.fabrics_intakes) != before:
			self._save_json_list(self.fabrics_intakes_file, self.fabrics_intakes)
			return True
		return False

	# ===== Unbarcoded Fabrics (בדים בלי ברקוד) =====
	def refresh_fabrics_unbarcoded(self):
		"""Reload unbarcoded fabrics from disk."""
		self.fabrics_unbarcoded = self._load_json_list(self.fabrics_unbarcoded_file)

	def add_unbarcoded_fabric(self, fabric_type: str, manufacturer: str = '', color: str = '', shade: str = '', notes: str = '') -> int:
		"""Add a single unbarcoded fabric row and return its ID."""
		try:
			if not (fabric_type or '').strip():
				raise Exception("חובה לציין 'סוג בד'")
			new_id = max([r.get('id', 0) for r in self.fabrics_unbarcoded], default=0) + 1
			rec = {
				'id': new_id,
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
				'fabric_type': (fabric_type or '').strip(),
				'manufacturer': (manufacturer or '').strip(),
				'color': (color or '').strip(),
				'shade': (shade or '').strip(),
				'notes': (notes or '').strip()
			}
			self.fabrics_unbarcoded.append(rec)
			self._save_json_list(self.fabrics_unbarcoded_file, self.fabrics_unbarcoded)
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת בד ללא ברקוד: {e}")

	def delete_unbarcoded_fabric(self, rec_id: int) -> bool:
		"""Delete unbarcoded fabric by ID."""
		before = len(self.fabrics_unbarcoded)
		try:
			self.fabrics_unbarcoded = [r for r in self.fabrics_unbarcoded if int(r.get('id', -1)) != int(rec_id)]
		except Exception:
			self.fabrics_unbarcoded = [r for r in self.fabrics_unbarcoded if (r.get('id') != rec_id)]
		if len(self.fabrics_unbarcoded) != before:
			self._save_json_list(self.fabrics_unbarcoded_file, self.fabrics_unbarcoded)
			return True
		return False
    
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
    
	def add_to_local_table(self, results: List[Dict], file_name: str = "", fabric_type: str = "", recipient_supplier: str = "", estimated_layers: int = 200, marker_width: float = None, marker_length: float = None) -> int:
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
			
			# הוספת כמות שכבות משוערת
			record['כמות שכבות משוערת'] = estimated_layers
			
			# הוספת נתוני מידות ציור אם קיימים
			if marker_width is not None:
				record['רוחב ציור'] = marker_width
			if marker_length is not None:
				record['אורך ציור'] = marker_length
            
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

	def update_drawing_weight_and_meters(self, drawing_id: int, total_weight: float, total_meters: float) -> bool:
		"""עדכון משקל ומטרים לציור שנחתך. מחזיר True אם בוצע שינוי ונשמר."""
		try:
			if total_weight is None and total_meters is None:
				return False
			changed = False
			for rec in self.drawings_data:
				if rec.get('id') == drawing_id:
					# עדכון משקל אם סופק
					if total_weight is not None:
						prev_weight = rec.get('משקל כולל')
						if prev_weight != total_weight:
							rec['משקל כולל'] = float(total_weight)
							changed = True
					# עדכון מטרים אם סופק
					if total_meters is not None:
						prev_meters = rec.get('מטרים כוללים')
						if prev_meters != total_meters:
							rec['מטרים כוללים'] = float(total_meters)
							changed = True
					break
			if changed:
				self.save_drawings_data()
			return changed
		except Exception as e:
			print(f"שגיאה בעדכון משקל ומטרים לציור: {e}")
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

	def bulk_update_fabrics(self, updates: list[dict]) -> int:
		"""עדכון מרובה של בדים לפי ברקודים.

		updates: [{'barcode': str, 'status': optional str, 'location': optional str}, ...]
		מחזיר מספר רשומות שעודכנו בפועל.
		"""
		try:
			if not updates:
				return 0
			upd_map = {}
			for u in updates:
				bc = str((u or {}).get('barcode', '')).strip()
				if not bc:
					continue
				upd_map[bc] = {'status': u.get('status'), 'location': u.get('location')}
			changed = 0
			for rec in self.fabrics_inventory:
				bc = str(rec.get('barcode', '')).strip()
				if not bc or bc not in upd_map:
					continue
				info = upd_map[bc]
				before = (rec.get('status'), rec.get('location'))
				if info.get('status') is not None:
					rec['status'] = info['status']
				if info.get('location') is not None:
					rec['location'] = info['location']
				after = (rec.get('status'), rec.get('location'))
				if after != before:
					changed += 1
			if changed:
				self.save_fabrics_inventory()
			return changed
		except Exception as e:
			print(f"שגיאה בעדכון מרובה של בדים: {e}")
			return 0

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
								 category: str = '', ticks_qty: int | str = 0, elastic_qty: int | str = 0, ribbon_qty: int | str = 0, fabric_category: str = '', square_area: float | str = 0.0) -> tuple[int, str]:
		"""הוספת מוצר לקטלוג עם שדות מורחבים. מחזיר (ID חדש, ברקוד חדש).

		:param name: שם מוצר (חובה)
		:param size: מידה
		:param fabric_type: סוג בד
		:param fabric_color: צבע בד
		:param print_name: שם פרינט
		:param category: קטגוריה (אופציונלי)
		:param ticks_qty: כמות טיקטקים (אופציונלי, מספר שלם)
		:param elastic_qty: כמות גומי (אופציונלי, מספר שלם)
		:param ribbon_qty: כמות סרט (אופציונלי, מספר שלם)
		:param fabric_category: קטגוריית בד (אופציונלי)
		:param square_area: שטח רבוע (אופציונלי, מספר עשרוני)
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
			def _to_float(val):
				try:
					if val in (None, ''): return 0.0
					return float(str(val).strip())
				except Exception:
					return 0.0
			ticks_i = _to_int(ticks_qty)
			elastic_i = _to_int(elastic_qty)
			ribbon_i = _to_int(ribbon_qty)
			square_area_f = _to_float(square_area)
			new_id = max([p.get('id', 0) for p in self.products_catalog], default=0) + 1
			
			# Generate new barcode
			last_barcode = self.barcodes_data.get('last_barcode', '7297555019592')
			new_barcode = self.generate_next_barcode(last_barcode)
			
			record = {
				'id': new_id,
				'barcode': new_barcode,
				'name': name.strip(),
				'size': size.strip(),
				'fabric_type': fabric_type.strip(),
				'fabric_color': fabric_color.strip(),
				'print_name': print_name.strip(),
				'category': (category or '').strip(),
				'fabric_category': (fabric_category or '').strip(),
				'square_area': square_area_f,
				'ticks_qty': ticks_i,
				'elastic_qty': elastic_i,
				'ribbon_qty': ribbon_i,
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			self.products_catalog.append(record)
			self.save_products_catalog()
			
			# Update last barcode
			self.save_barcodes_data(new_barcode)
			
			return (new_id, new_barcode)
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
			# בונים DataFrame מסודר עם עמודות קבועות, כולל barcode, fabric_category ו-square_area
			columns = [
				'barcode','id','name','category','size','fabric_type','fabric_color','fabric_category',
				'print_name','square_area','ticks_qty','elastic_qty','ribbon_qty','created_at'
			]
			rows = []
			for rec in self.products_catalog:
				rows.append({
					'barcode': rec.get('barcode',''),
					'id': rec.get('id'),
					'name': rec.get('name',''),
					'category': rec.get('category',''),
					'size': rec.get('size',''),
					'fabric_type': rec.get('fabric_type',''),
					'fabric_color': rec.get('fabric_color',''),
					# ברירת מחדל: "בלי קטגוריה" אם חסר
					'fabric_category': rec.get('fabric_category') or 'בלי קטגוריה',
					'print_name': rec.get('print_name',''),
					'square_area': rec.get('square_area', 0.0),
					'ticks_qty': rec.get('ticks_qty', 0),
					'elastic_qty': rec.get('elastic_qty', 0),
					'ribbon_qty': rec.get('ribbon_qty', 0),
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
			# נוודא קיום עמודות נדרשות
			required = {'name','category','size','fabric_type','fabric_color','fabric_category','print_name','square_area','ticks_qty','elastic_qty','ribbon_qty','created_at'}
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
			
			def _to_float(x):
				try:
					if x in (None, ''):
						return 0.0
					return float(x)
				except Exception:
					return 0.0

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
					square_area = _to_float(row.get('square_area'))
					ticks = _to_int(row.get('ticks_qty'))
					elastic = _to_int(row.get('elastic_qty'))
					ribbon = _to_int(row.get('ribbon_qty'))
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
						'square_area': square_area,
						'ticks_qty': ticks,
						'elastic_qty': elastic,
						'ribbon_qty': ribbon,
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
					square_area = _to_float(row.get('square_area'))
					ticks = _to_int(row.get('ticks_qty'))
					elastic = _to_int(row.get('elastic_qty'))
					ribbon = _to_int(row.get('ribbon_qty'))
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
						'square_area': square_area,
						'ticks_qty': ticks,
						'elastic_qty': elastic,
						'ribbon_qty': ribbon,
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

	def _add_to_simple_list(self, data_list: list[dict], save_func, name: str) -> int:
		if not name:
			raise Exception("חובה להזין שם")
		for rec in data_list:
			if rec.get('name','').strip() == name.strip():
				raise Exception("פריט כבר קיים")
		new_id = max([r.get('id',0) for r in data_list], default=0) + 1
		rec = {'id': new_id, 'name': name.strip(), 'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
		data_list.append(rec)
		save_func(); return new_id

	def _delete_from_simple_list(self, data_list: list[dict], save_func, rec_id: int) -> bool:
		before = len(data_list)
		data_list[:] = [r for r in data_list if int(r.get('id',0)) != int(rec_id)]
		if len(data_list) != before:
			save_func(); return True
		return False

	def add_product_size(self, name: str) -> int:
		return self._add_to_simple_list(self.product_sizes, self.save_product_sizes, name)

	def delete_product_size(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_sizes, self.save_product_sizes, rec_id)

	def add_fabric_type_item(self, name: str) -> int:
		return self._add_to_simple_list(self.product_fabric_types, self.save_fabric_types, name)

	def delete_fabric_type_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_types, self.save_fabric_types, rec_id)

	def add_fabric_color_item(self, name: str) -> int:
		return self._add_to_simple_list(self.product_fabric_colors, self.save_fabric_colors, name)

	def delete_fabric_color_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_colors, self.save_fabric_colors, rec_id)

	def add_print_name_item(self, name: str) -> int:
		return self._add_to_simple_list(self.product_print_names, self.save_print_names, name)

	def delete_print_name_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_print_names, self.save_print_names, rec_id)

	def add_fabric_category_item(self, name: str) -> int:
		return self._add_to_simple_list(self.product_fabric_categories, self.save_fabric_categories, name)

	def delete_fabric_category_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_fabric_categories, self.save_fabric_categories, rec_id)

	def add_model_name_item(self, name: str, sewing_price: float = 0.0) -> int:
		"""הוספת שם דגם עם מחיר תפירה"""
		if not name:
			raise Exception("חובה להזין שם")
		for rec in self.product_model_names:
			if rec.get('name','').strip() == name.strip():
				raise Exception("פריט כבר קיים")
		new_id = max([r.get('id',0) for r in self.product_model_names], default=0) + 1
		rec = {
			'id': new_id,
			'name': name.strip(),
			'sewing_price': float(sewing_price),
			'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
		}
		self.product_model_names.append(rec)
		self.save_model_names()
		return new_id

	def update_model_name_sewing_price(self, rec_id: int, sewing_price: float) -> bool:
		"""עדכון מחיר תפירה לשם דגם"""
		try:
			for rec in self.product_model_names:
				if rec.get('id') == rec_id:
					rec['sewing_price'] = float(sewing_price)
					self.save_model_names()
					return True
			return False
		except Exception as e:
			print(f"שגיאה בעדכון מחיר תפירה: {e}")
			return False

	def get_sewing_price_by_model_name(self, model_name: str) -> float:
		"""קבלת מחיר תפירה לפי שם דגם"""
		for rec in self.product_model_names:
			if rec.get('name', '').strip() == model_name.strip():
				return float(rec.get('sewing_price', 0) or 0)
		return 0.0

	def delete_model_name_item(self, rec_id: int) -> bool:
		return self._delete_from_simple_list(self.product_model_names, self.save_model_names, rec_id)

	def refresh_product_attributes(self):
		self.product_sizes = self.load_product_sizes()
		self.product_fabric_types = self.load_fabric_types()
		self.product_fabric_colors = self.load_fabric_colors()
		self.product_print_names = self.load_print_names()
		self.product_fabric_categories = self.load_fabric_categories()
		self.product_model_names = self.load_model_names()

	# ===== Fabric Prices (מחירי בדים) =====
	def load_fabric_prices(self) -> List[Dict]:
		"""טעינת טבלת מחירי בדים"""
		try:
			path = 'fabric_prices.json'
			if os.path.exists(path):
				with open(path, 'r', encoding='utf-8') as f:
					return json.load(f)
			return []
		except Exception as e:
			print(f"שגיאה בטעינת מחירי בדים: {e}")
			return []

	def save_fabric_prices(self, prices: List[Dict]) -> bool:
		"""שמירת טבלת מחירי בדים"""
		try:
			with open('fabric_prices.json', 'w', encoding='utf-8') as f:
				json.dump(prices, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת מחירי בדים: {e}")
			return False

	def add_fabric_price(self, fabric_category: str, fabric_color: str, print_name: str, price_per_sqm: float) -> int:
		"""הוספת מחיר בד חדש"""
		try:
			prices = self.load_fabric_prices()
			# בדיקת כפילות
			for p in prices:
				if (p.get('fabric_category', '').strip() == fabric_category.strip() and
					p.get('fabric_color', '').strip() == fabric_color.strip() and
					p.get('print_name', '').strip() == print_name.strip()):
					raise ValueError("שילוב זה כבר קיים בטבלה")
			new_id = max([p.get('id', 0) for p in prices], default=0) + 1
			rec = {
				'id': new_id,
				'fabric_category': fabric_category.strip(),
				'fabric_color': fabric_color.strip(),
				'print_name': print_name.strip(),
				'price_per_sqm': float(price_per_sqm),
				'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
			}
			prices.append(rec)
			self.save_fabric_prices(prices)
			return new_id
		except Exception as e:
			raise Exception(f"שגיאה בהוספת מחיר בד: {str(e)}")

	def update_fabric_price(self, price_id: int, fabric_category: str, fabric_color: str, print_name: str, price_per_sqm: float) -> bool:
		"""עדכון מחיר בד קיים"""
		try:
			prices = self.load_fabric_prices()
			for p in prices:
				if p.get('id') == price_id:
					p['fabric_category'] = fabric_category.strip()
					p['fabric_color'] = fabric_color.strip()
					p['print_name'] = print_name.strip()
					p['price_per_sqm'] = float(price_per_sqm)
					p['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
					return self.save_fabric_prices(prices)
			return False
		except Exception as e:
			print(f"שגיאה בעדכון מחיר בד: {e}")
			return False

	def delete_fabric_price(self, price_id: int) -> bool:
		"""מחיקת מחיר בד"""
		try:
			prices = self.load_fabric_prices()
			before = len(prices)
			prices = [p for p in prices if p.get('id') != price_id]
			if len(prices) != before:
				return self.save_fabric_prices(prices)
			return False
		except Exception as e:
			print(f"שגיאה במחיקת מחיר בד: {e}")
			return False

	def find_fabric_price(self, fabric_category: str, fabric_color: str, print_name: str) -> Dict:
		"""מציאת מחיר בד לפי קטגוריה, צבע ופרינט"""
		prices = self.load_fabric_prices()
		for p in prices:
			if (p.get('fabric_category', '').strip() == fabric_category.strip() and
				p.get('fabric_color', '').strip() == fabric_color.strip() and
				p.get('print_name', '').strip() == print_name.strip()):
				return p
		return {}

	# ===== Global Item Cost Settings (הגדרות עלויות גלובליות) =====
	def load_item_cost_settings(self) -> Dict:
		"""טעינת הגדרות עלויות פריטים"""
		try:
			path = 'item_cost_settings.json'
			if os.path.exists(path):
				with open(path, 'r', encoding='utf-8') as f:
					return json.load(f)
			# ברירת מחדל
			return {
				'tick_price': 0.0,
				'elastic_price': 0.0,
				'ribbon_price': 0.0,
				'sewing_price': 0.0
			}
		except Exception as e:
			print(f"שגיאה בטעינת הגדרות עלויות: {e}")
			return {'tick_price': 0.0, 'elastic_price': 0.0, 'ribbon_price': 0.0, 'sewing_price': 0.0}

	def save_item_cost_settings(self, settings: Dict) -> bool:
		"""שמירת הגדרות עלויות פריטים"""
		try:
			with open('item_cost_settings.json', 'w', encoding='utf-8') as f:
				json.dump(settings, f, indent=2, ensure_ascii=False)
			return True
		except Exception as e:
			print(f"שגיאה בשמירת הגדרות עלויות: {e}")
			return False

	def calculate_item_cost(self, item: Dict) -> Dict:
		"""חישוב עלות פריט"""
		try:
			settings = self.load_item_cost_settings()
			# מציאת מחיר בד
			fabric_price_rec = self.find_fabric_price(
				item.get('fabric_category', ''),
				item.get('fabric_color', ''),
				item.get('print_name', '')
			)
			
			square_area = float(item.get('square_area', 0) or 0)
			ticks_qty = float(item.get('ticks_qty', 0) or 0)
			elastic_qty = float(item.get('elastic_qty', 0) or 0)
			ribbon_qty = float(item.get('ribbon_qty', 0) or 0)
			
			# חישוב עלות בד: שטח_רבוע × מחיר_למ"ר
			price_per_sqm = float(fabric_price_rec.get('price_per_sqm', 0) or 0)
			fabric_cost = square_area * price_per_sqm
			
			# חישוב עלויות אביזרים
			tick_price = float(settings.get('tick_price', 0) or 0)
			elastic_price = float(settings.get('elastic_price', 0) or 0)
			ribbon_price = float(settings.get('ribbon_price', 0) or 0)
			
			# מחיר תפירה - לפי שם הדגם (או גלובלי אם לא מוגדר)
			model_name = item.get('name', '')
			sewing_price = self.get_sewing_price_by_model_name(model_name)
			if sewing_price == 0:
				sewing_price = float(settings.get('sewing_price', 0) or 0)
			
			ticks_cost = ticks_qty * tick_price
			elastic_cost = elastic_qty * elastic_price
			ribbon_cost = ribbon_qty * ribbon_price
			
			total_cost = fabric_cost + ticks_cost + elastic_cost + ribbon_cost + sewing_price
			
			return {
				'fabric_cost': round(fabric_cost, 2),
				'ticks_cost': round(ticks_cost, 2),
				'elastic_cost': round(elastic_cost, 2),
				'ribbon_cost': round(ribbon_cost, 2),
				'sewing_cost': round(sewing_price, 2),
				'total_cost': round(total_cost, 2)
			}
		except Exception as e:
			print(f"שגיאה בחישוב עלות פריט: {e}")
			return {
				'fabric_cost': 0,
				'ticks_cost': 0,
				'elastic_cost': 0,
				'ribbon_cost': 0,
				'sewing_cost': 0,
				'total_cost': 0
			}
