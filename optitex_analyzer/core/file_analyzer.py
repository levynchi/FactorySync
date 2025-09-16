"""
ניתוח קבצי אופטיטקס ועיבוד נתונים
"""

import pandas as pd
import os
import re
from typing import Dict, List, Tuple, Optional

class OptitexFileAnalyzer:
    """מנתח קבצי אופטיטקס"""
    
    def __init__(self):
        self.product_mapping = {}
        self.is_tubular = False
        self.results = []
        # Marker meta
        self.marker_width = None
        self.marker_length = None
        # Exceptions: style files that should NOT be treated as tubular even if layout reports Tubular
        self._tubular_exceptions = {
            # exact base filename match (case-insensitive compare will be used)
            '200 baby leggings close.pds'
        }
        
    def load_products_mapping(self, products_file: str) -> bool:
        """טעינת מיפוי מוצרים מקובץ Excel"""
        try:
            df_products = pd.read_excel(products_file)
            self.product_mapping = {}
            
            for _, row in df_products.iterrows():
                if pd.notna(row['product name']):
                    self.product_mapping[row['file name']] = row['product name']
            
            return True
        except Exception as e:
            raise Exception(f"שגיאה בטעינת קובץ מוצרים: {str(e)}")
    
    def analyze_file(self, rib_file: str, handle_tubular: bool = True, 
                    only_positive: bool = True) -> List[Dict]:
        """ניתוח קובץ אופטיטקס"""
        try:
            # קריאת הקובץ
            df = pd.read_excel(rib_file, header=None)
            # איפוס נתוני מרקר
            self.marker_width = None
            self.marker_length = None
            
            # בדיקת Tubular (ברמת הקובץ)
            self.is_tubular = False
            if handle_tubular:
                self.is_tubular = self._check_tubular_layout(df)
            
            # חיפוש נתונים
            self.results = []
            current_file_name = None
            current_style_tubular = self.is_tubular  # ברירת מחדל: כמו רמת הקובץ
            product_name = None
            
            for i, row in df.iterrows():
                # ===== חילוץ Marker Width / Length משורה =====
                # כדי למנוע NaN נשתמש ברג'קס לזיהוי מספרים גם אם יש טקסט (למשל "180 cm" או "Width: 180")
                try:
                    if (self.marker_width is None) or (self.marker_length is None):
                        # המרה לרשימת תאים (שומרים את האובייקט המקורי לצורך מציאת המספר)
                        row_cells = list(row)
                        # ניצור רשימת מחרוזות נורמליזציה לחיפוש טקסטואלי
                        normalized = [str(c).strip().lower() if (isinstance(c, str) or not pd.isna(c)) else '' for c in row_cells]

                        def parse_numeric(val):
                            if pd.isna(val):
                                return None
                            if isinstance(val, (int, float)):
                                return float(val)
                            s = str(val).strip()
                            # החלפת פסיקים בנקודות (פורמט אירופאי)
                            s = s.replace(',', '.')
                            m = re.search(r'[-+]?\d+(?:\.\d+)?', s)
                            if m:
                                try:
                                    return float(m.group())
                                except:
                                    return None
                            return None

                        def find_number_near(index: int) -> Optional[float]:
                            # עמודה קודמת, הבאה, ראשונה, כל השורה
                            candidate_indices = []
                            if index > 0:
                                candidate_indices.append(index - 1)
                            if index + 1 < len(row_cells):
                                candidate_indices.append(index + 1)
                            candidate_indices.append(0)
                            candidate_indices.extend(range(len(row_cells)))  # fallback חיפוש כללי
                            seen = set()
                            for ci in candidate_indices:
                                if ci in seen:
                                    continue
                                seen.add(ci)
                                val = parse_numeric(row_cells[ci])
                                if val is not None:
                                    return val
                            return None

                        for col_idx, text in enumerate(normalized):
                            if (self.marker_width is not None) and (self.marker_length is not None):
                                break
                            if 'marker' in text and ('width' in text or 'wid' in text):
                                if self.marker_width is None:
                                    num = find_number_near(col_idx)
                                    if num is not None:
                                        self.marker_width = num
                            elif 'marker' in text and ('length' in text or 'len' in text):
                                if self.marker_length is None:
                                    num = find_number_near(col_idx)
                                    if num is not None:
                                        self.marker_length = num
                except Exception:
                    # לא נכשיל את הניתוח בגלל שגיאת חילוץ מרקר
                    pass
                # חיפוש שם קובץ
                if pd.notna(row.iloc[0]) and row.iloc[0] == 'Style File Name:':
                    if pd.notna(row.iloc[2]):
                        current_file_name = os.path.basename(row.iloc[2])
                        # קבע טיפול Tubular ברמת ה-Style: אם בקובץ Tubular אך ה-Style חריג, בטל חלוקה רק עבור Style זה
                        current_style_tubular = self.is_tubular
                        if current_file_name and isinstance(current_file_name, str):
                            if current_file_name.strip().lower() in self._tubular_exceptions:
                                current_style_tubular = False
                        product_name = self.product_mapping.get(current_file_name)

                # חיפוש טבלת מידות
                elif (pd.notna(row.iloc[0]) and row.iloc[0] == 'Size name' and 
                      pd.notna(row.iloc[1]) and row.iloc[1] == 'Order' and product_name):
                    # עיבוד טבלת מידות עם אינדיקציה האם לחלק ב-2 עבור Style זה
                    self._process_sizes_table(df, i, product_name, only_positive, apply_tubular=current_style_tubular)
            
            return self.results
            
        except Exception as e:
            raise Exception(f"שגיאה בניתוח הקובץ: {str(e)}")
    
    def _check_tubular_layout(self, df: pd.DataFrame) -> bool:
        """בדיקה אם יש Layout Tubular"""
        for i, row in df.iterrows():
            if (pd.notna(row.iloc[0]) and row.iloc[0] == 'Layout' and 
                pd.notna(row.iloc[2]) and row.iloc[2] == 'Tubular'):
                return True
        return False
    
    def _clean_size_name(self, size_name: str) -> str:
        """ניקוי שם מידה - הסרת האות 'M' ממידות"""
        if not isinstance(size_name, str):
            return str(size_name)
        
        # הסרת האות 'M' ממידות בפורמט כמו "3M-6M" או "6M-12M"
        cleaned = size_name.replace('M', '')
        
        # ניקוי נוסף של רווחים מיותרים
        cleaned = cleaned.strip()
        
        return cleaned
    
    def _process_sizes_table(self, df: pd.DataFrame, start_index: int, 
                           product_name: str, only_positive: bool, apply_tubular: bool):
        """עיבוד טבלת מידות"""
        j = start_index + 1
        
        while j < len(df) and pd.notna(df.iloc[j, 0]) and pd.notna(df.iloc[j, 1]):
            size_name = df.iloc[j, 0]
            
            try:
                quantity = int(df.iloc[j, 1])
            except (ValueError, TypeError):
                j += 1
                continue
            
            if size_name not in ['Style File Name:', 'Size name']:
                # ניקוי מידה - הסרת האות 'M' ממידות
                cleaned_size = self._clean_size_name(size_name)
                
                # טיפול ב-Tubular
                original_quantity = quantity
                if apply_tubular and quantity > 0:
                    quantity = quantity / 2
                    quantity = int(quantity) if quantity == int(quantity) else round(quantity, 1)
                
                # הוספה לתוצאות
                if not only_positive or quantity > 0:
                    self.results.append({
                        'שם המוצר': product_name,
                        'מידה': cleaned_size,
                        'כמות': quantity,
                        'כמות מקורית': original_quantity if apply_tubular else quantity,
                        'הערה': 'חולק ב-2 (Tubular)' if apply_tubular and original_quantity > 0 else 'רגיל'
                    })
            elif size_name in ['Style File Name:', 'Size name']:
                break
            
            j += 1
    
    def get_analysis_summary(self) -> Dict:
        """קבלת סיכום הניתוח"""
        if not self.results:
            return {}
        
        df = pd.DataFrame(self.results)
        
        return {
            'total_records': len(self.results),
            'unique_products': df['שם המוצר'].nunique(),
            'unique_sizes': df['מידה'].nunique(),
            'total_quantity': df['כמות'].sum(),
            'is_tubular': self.is_tubular,
            'products_list': df['שם המוצר'].unique().tolist(),
            'marker_width': self.marker_width,
            'marker_length': self.marker_length
        }
    
    def sort_results(self) -> List[Dict]:
        """מיון התוצאות לפי מוצר ומידה"""
        if not self.results:
            return []
        
        df = pd.DataFrame(self.results)
        
        def sort_size(size):
            """פונקציה למיון מידות"""
            size_str = str(size)
            
            # טיפול במידות חודשים (0-3, 3-6, וכו')
            if '-' in size_str and any(char.isdigit() for char in size_str):
                try:
                    first_num = size_str.split('-')[0]
                    if first_num.isdigit():
                        return (0, int(first_num))
                except:
                    pass
            
            try:
                # מידות נומריות
                return (1, float(size))
            except:
                # מידות טקסטואליות
                return (2, str(size))
        
        # מיון
        df['_sort_size'] = df['מידה'].apply(sort_size)
        df = df.sort_values(['שם המוצר', '_sort_size'])
        df = df.drop('_sort_size', axis=1)
        
        return df.to_dict('records')
    
    def get_products_found(self) -> List[Tuple[str, str]]:
        """קבלת רשימת המוצרים שנמצאו"""
        found_products = []
        processed_files = set()
        
        for result in self.results:
            product_name = result['שם המוצר']
            # חיפוש שם הקובץ המקורי
            for file_name, mapped_name in self.product_mapping.items():
                if mapped_name == product_name and file_name not in processed_files:
                    found_products.append((file_name, product_name))
                    processed_files.add(file_name)
                    break
        
        return found_products
