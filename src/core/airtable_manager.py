"""
מודול לניהול חיבור ל-Airtable
"""

import os
from pyairtable import Api


class AirtableManager:
    """מחלקה לניהול חיבור ל-Airtable"""
    
    def __init__(self, api_key="", base_id="", table_id=""):
        self.api_key = api_key
        self.base_id = base_id
        self.table_id = table_id
        self.api = None
        self.table = None
    
    def connect(self):
        """חיבור ל-Airtable"""
        try:
            if not self.api_key or not self.base_id:
                return {
                    'success': False,
                    'error': 'חסרים פרטי חיבור (API Key או Base ID)'
                }
            
            self.api = Api(self.api_key)
            self.table = self.api.table(self.base_id, self.table_id)
            
            return {
                'success': True,
                'message': 'חיבור ל-Airtable הצליח'
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f'שגיאה בחיבור ל-Airtable: {str(e)}'
            }
    
    def upload_drawing(self, file_name, output_df):
        """העלאת ציור ל-Airtable"""
        try:
            # חיבור ל-Airtable
            connection_result = self.connect()
            if not connection_result['success']:
                return connection_result
            
            # הכנת הנתונים
            airtable_record = {
                'שם הקובץ': os.path.splitext(os.path.basename(file_name))[0] if file_name else '',
            }
            
            # הוספת הדגמים והכמויות לעמודות נפרדות
            model_index = 1
            for _, row in output_df.iterrows():
                if model_index > 21:  # מקסימום 21 דגמים באייר טייבל
                    break
                    
                product_name = row['שם המוצר']
                size = row['מידה']
                quantity = row['כמות']
                
                # יצירת מחרוזת דגם משולבת
                model_name = f"{product_name} {size}"
                
                # הוספה לעמודות
                airtable_record[f'דגם {model_index}'] = model_name
                airtable_record[f'כמות דגם {model_index}'] = int(quantity) if quantity == int(quantity) else quantity
                
                model_index += 1
            
            # העלאת הרשומה
            created_record = self.table.create(airtable_record)
            
            return {
                'success': True,
                'record_id': created_record['id'],
                'models_count': model_index - 1,
                'record_data': airtable_record
            }
            
        except Exception as e:
            error_msg = str(e)
            
            # הצעת פתרונות לבעיות נפוצות
            help_message = ""
            if "Invalid API key" in error_msg:
                help_message = "API Key לא תקין. בדוק שהמפתח נכון ופעיל."
            elif "Base not found" in error_msg:
                help_message = "Base ID לא נמצא. בדוק שה-Base ID נכון."
            elif "Table not found" in error_msg:
                help_message = "הטבלה לא נמצאה. בדוק שה-Table ID נכון."
            
            return {
                'success': False,
                'error': error_msg,
                'help': help_message
            }
    
    def test_connection(self):
        """בדיקת חיבור ל-Airtable"""
        try:
            connection_result = self.connect()
            if not connection_result['success']:
                return connection_result
            
            # נסיון לקרוא מהטבלה (בדיקה בסיסית)
            records = self.table.all(max_records=1)
            
            return {
                'success': True,
                'message': 'החיבור ל-Airtable תקין',
                'table_accessible': True
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f'שגיאה בבדיקת חיבור: {str(e)}',
                'table_accessible': False
            }
