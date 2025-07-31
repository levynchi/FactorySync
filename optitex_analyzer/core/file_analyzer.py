"""
ניתוח קבצי אופטיטקס ועיבוד נתונים
"""

import pandas as pd
import os
from typing import Dict, List, Tuple, Optional

class OptitexFileAnalyzer:
    """מנתח קבצי אופטיטקס"""
    
    def __init__(self):
        self.product_mapping = {}
        self.is_tubular = False
        self.results = []
    
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
            
            # בדיקת Tubular
            self.is_tubular = False
            if handle_tubular:
                self.is_tubular = self._check_tubular_layout(df)
            
            # חיפוש נתונים
            self.results = []
            current_file_name = None
            product_name = None
            
            for i, row in df.iterrows():
                # חיפוש שם קובץ
                if pd.notna(row.iloc[0]) and row.iloc[0] == 'Style File Name:':
                    if pd.notna(row.iloc[2]):
                        current_file_name = os.path.basename(row.iloc[2])
                        product_name = self.product_mapping.get(current_file_name)
                
                # חיפוש טבלת מידות
                elif (pd.notna(row.iloc[0]) and row.iloc[0] == 'Size name' and 
                      pd.notna(row.iloc[1]) and row.iloc[1] == 'Order' and product_name):
                    
                    self._process_sizes_table(df, i, product_name, only_positive)
            
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
    
    def _process_sizes_table(self, df: pd.DataFrame, start_index: int, 
                           product_name: str, only_positive: bool):
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
                # טיפול ב-Tubular
                original_quantity = quantity
                if self.is_tubular and quantity > 0:
                    quantity = quantity / 2
                    quantity = int(quantity) if quantity == int(quantity) else round(quantity, 1)
                
                # הוספה לתוצאות
                if not only_positive or quantity > 0:
                    self.results.append({
                        'שם המוצר': product_name,
                        'מידה': size_name,
                        'כמות': quantity,
                        'כמות מקורית': original_quantity if self.is_tubular else quantity,
                        'הערה': 'חולק ב-2 (Tubular)' if self.is_tubular and original_quantity > 0 else 'רגיל'
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
            'products_list': df['שם המוצר'].unique().tolist()
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
