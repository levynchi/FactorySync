"""
מודול לניתוח קבצי אופטיטקס והמרתם לפורמט נקי
"""

import pandas as pd
import os


class OptitexAnalyzer:
    """מחלקה לניתוח קבצי אופטיטקס"""
    
    def __init__(self):
        self.rib_file = ""
        self.products_file = ""
        self.output_df = None
    
    @staticmethod
    def load_product_mapping(products_file):
        """טעינת מיפוי מוצרים מקובץ Excel"""
        product_mapping = {}
        try:
            df_products = pd.read_excel(products_file)
            for _, row in df_products.iterrows():
                if pd.notna(row['product name']):
                    product_mapping[row['file name']] = row['product name']
        except Exception as e:
            raise Exception(f"שגיאה בטעינת קובץ מוצרים: {str(e)}")
        
        return product_mapping
    
    @staticmethod
    def check_tubular_layout(df):
        """בדיקה אם הקובץ מכיל Layout Tubular"""
        for i, row in df.iterrows():
            if (pd.notna(row.iloc[0]) and row.iloc[0] == 'Layout' and 
                pd.notna(row.iloc[2]) and row.iloc[2] == 'Tubular'):
                return True
        return False
    
    @staticmethod
    def sort_size(size):
        """פונקציה למיון מידות בסדר נומרי/אלפביתי נכון"""
        size_str = str(size)
        
        # טיפול במידות חודשים (0-3, 3-6, 6-12, 12-18, 18-24, 24-30)
        if '-' in size_str and any(char.isdigit() for char in size_str):
            try:
                # חילוץ המספר הראשון מהמידה
                first_num = size_str.split('-')[0]
                if first_num.isdigit():
                    return (0, int(first_num))  # מידות חודשים - קבוצה 0
            except:
                pass
        
        try:
            # אם המידה היא מספר, נמיין לפי ערך נומרי
            return (1, float(size))  # מידות נומריות רגילות - קבוצה 1
        except:
            # אם המידה היא טקסט, נמיין לפי אלפבית
            return (2, str(size))  # מידות טקסטואליות - קבוצה 2
    
    def analyze_files(self, rib_file, products_file, check_tubular=True, only_positive=True):
        """ניתוח קבצי RIB ומוצרים"""
        self.rib_file = rib_file
        self.products_file = products_file
        
        try:
            # טעינת מיפוי מוצרים
            product_mapping = self.load_product_mapping(products_file)
            
            # קריאת קובץ RIB
            df = pd.read_excel(rib_file, header=None)
            
            # בדיקת Tubular
            is_tubular = False
            if check_tubular:
                is_tubular = self.check_tubular_layout(df)
            
            # חיפוש נתונים
            results = []
            current_file_name = None
            product_name = None
            
            for i, row in df.iterrows():
                # חיפוש שם קובץ
                if pd.notna(row.iloc[0]) and row.iloc[0] == 'Style File Name:':
                    if pd.notna(row.iloc[2]):
                        current_file_name = os.path.basename(row.iloc[2])
                        product_name = product_mapping.get(current_file_name)
                
                # חיפוש טבלת מידות
                elif (pd.notna(row.iloc[0]) and row.iloc[0] == 'Size name' and 
                      pd.notna(row.iloc[1]) and row.iloc[1] == 'Order' and product_name):
                    
                    j = i + 1
                    while j < len(df) and pd.notna(df.iloc[j, 0]) and pd.notna(df.iloc[j, 1]):
                        size_name = df.iloc[j, 0]
                        quantity = int(df.iloc[j, 1])
                        
                        if size_name not in ['Style File Name:', 'Size name']:
                            # טיפול ב-Tubular
                            original_quantity = quantity
                            if is_tubular and quantity > 0:
                                quantity = quantity / 2
                                quantity = int(quantity) if quantity == int(quantity) else round(quantity, 1)
                            
                            # הוספה לתוצאות
                            if not only_positive or quantity > 0:
                                results.append({
                                    'שם המוצר': product_name,
                                    'מידה': size_name,
                                    'כמות': quantity,
                                    'כמות מקורית': original_quantity if is_tubular else quantity,
                                    'הערה': 'חולק ב-2 (Tubular)' if is_tubular and original_quantity > 0 else 'רגיל'
                                })
                        elif size_name in ['Style File Name:', 'Size name']:
                            break
                        j += 1
            
            # יצירת טבלה
            if results:
                self.output_df = pd.DataFrame(results)
                
                # יצירת עמודה זמנית למיון
                self.output_df['_sort_size'] = self.output_df['מידה'].apply(self.sort_size)
                
                # מיון לפי שם מוצר ואז לפי מידה
                self.output_df = self.output_df.sort_values(['שם המוצר', '_sort_size'])
                
                # הסרת עמודת המיון הזמנית
                self.output_df = self.output_df.drop('_sort_size', axis=1)
                
                return {
                    'success': True,
                    'data': self.output_df,
                    'is_tubular': is_tubular,
                    'total_products': self.output_df['שם המוצר'].nunique(),
                    'total_sizes': self.output_df['מידה'].nunique(),
                    'total_quantity': self.output_df['כמות'].sum(),
                    'products_mapping_count': len(product_mapping)
                }
            else:
                return {
                    'success': False,
                    'error': 'לא נמצאו נתונים',
                    'products_mapping_count': len(product_mapping)
                }
                
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def save_to_excel(self, file_path):
        """שמירת התוצאות לקובץ Excel"""
        if self.output_df is None or self.output_df.empty:
            raise Exception("אין נתונים לשמירה")
        
        try:
            self.output_df.to_excel(file_path, index=False)
            return True
        except Exception as e:
            raise Exception(f"שגיאה בשמירה: {str(e)}")
