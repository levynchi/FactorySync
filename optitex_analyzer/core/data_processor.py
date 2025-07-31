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
    
    def __init__(self, drawings_file: str = "drawings_data.json"):
        self.drawings_file = drawings_file
        self.drawings_data = self.load_drawings_data()
    
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
    
    def add_to_local_table(self, results: List[Dict], file_name: str = "") -> int:
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
                'סך כמויות': 0
            }
            
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
