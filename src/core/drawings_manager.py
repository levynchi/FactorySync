"""
מודול לניהול נתוני הציורים המקומיים
"""

import json
import os
import pandas as pd
from datetime import datetime


class DrawingsManager:
    """מחלקה לניהול נתוני ציורים מקומיים"""
    
    def __init__(self, data_file="drawings_data.json"):
        self.data_file = data_file
        self.drawings_data = []
        self.load_data()
    
    def load_data(self):
        """טעינת נתוני ציורים מקובץ JSON"""
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.drawings_data = json.load(f)
        except Exception as e:
            print(f"שגיאה בטעינת נתוני ציורים: {e}")
            self.drawings_data = []
    
    def save_data(self):
        """שמירת נתוני ציורים לקובץ JSON"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.drawings_data, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"שגיאה בשמירת נתוני ציורים: {e}")
            return False
    
    def add_drawing(self, file_name, output_df):
        """הוספת ציור חדש לטבלה"""
        try:
            # יצירת רשומה חדשה
            record = {
                'id': len(self.drawings_data) + 1,
                'שם הקובץ': os.path.splitext(os.path.basename(file_name))[0] if file_name else 'לא ידוע',
                'תאריך יצירה': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'מוצרים': [],
                'סך כמויות': 0
            }
            
            # הוספת מידע על המוצרים
            for product_name in output_df['שם המוצר'].unique():
                product_data = output_df[output_df['שם המוצר'] == product_name]
                
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
            success = self.save_data()
            
            if success:
                return {
                    'success': True,
                    'record': record
                }
            else:
                return {
                    'success': False,
                    'error': 'שגיאה בשמירה'
                }
                
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def get_all_drawings(self):
        """קבלת כל הציורים"""
        return self.drawings_data
    
    def get_drawing_by_id(self, drawing_id):
        """קבלת ציור לפי ID"""
        for record in self.drawings_data:
            if record.get('id') == drawing_id:
                return record
        return None
    
    def delete_drawing(self, drawing_id):
        """מחיקת ציור לפי ID"""
        try:
            self.drawings_data = [r for r in self.drawings_data if r.get('id') != drawing_id]
            success = self.save_data()
            return {
                'success': success,
                'error': None if success else 'שגיאה בשמירה'
            }
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def clear_all_drawings(self):
        """מחיקת כל הציורים"""
        try:
            self.drawings_data = []
            success = self.save_data()
            return {
                'success': success,
                'error': None if success else 'שגיאה בשמירה'
            }
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def export_to_excel(self, file_path):
        """ייצוא כל הציורים לקובץ Excel"""
        try:
            if not self.drawings_data:
                raise Exception("אין ציורים לייצוא")
            
            # יצירת רשימה של כל השורות
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
                raise Exception("אין נתונים לייצוא")
            
            # יצירת DataFrame ושמירה
            df = pd.DataFrame(rows)
            df.to_excel(file_path, index=False)
            
            return {
                'success': True,
                'rows_count': len(rows)
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def get_statistics(self):
        """קבלת סטטיסטיקות על הציורים"""
        if not self.drawings_data:
            return {
                'total_drawings': 0,
                'total_products': 0,
                'total_quantity': 0
            }
        
        total_products = 0
        total_quantity = 0
        
        for record in self.drawings_data:
            total_products += len(record.get('מוצרים', []))
            total_quantity += record.get('סך כמויות', 0)
        
        return {
            'total_drawings': len(self.drawings_data),
            'total_products': total_products,
            'total_quantity': total_quantity
        }
