"""
חיבור ועבודה עם אייר טייבל
"""

from typing import Dict, List, Any
from pyairtable import Api

class AirtableClient:
    """לקוח אייר טייבל"""
    
    def __init__(self, api_key: str, base_id: str, table_id: str):
        self.api_key = api_key.strip()
        self.base_id = base_id.strip()
        self.table_id = table_id.strip()
        self.api = None
        self.table = None
        
        if self.api_key and self.base_id:
            self._initialize_connection()
    
    def _initialize_connection(self):
        """יצירת חיבור לאייר טייבל"""
        try:
            self.api = Api(self.api_key)
            self.table = self.api.table(self.base_id, self.table_id)
        except Exception as e:
            raise Exception(f"שגיאה ביצירת חיבור לאייר טייבל: {str(e)}")
    
    def test_connection(self) -> bool:
        """בדיקת חיבור לאייר טייבל"""
        try:
            if not self.table:
                return False
            
            # ניסיון לקרוא רשומה אחת כדי לבדוק את החיבור
            self.table.all(max_records=1)
            return True
        except Exception:
            return False
    
    def upload_record(self, record_data: Dict[str, Any]) -> Dict:
        """העלאת רשומה לאייר טייבל"""
        try:
            if not self.table:
                raise Exception("חיבור לאייר טייבל לא הוקם")
            
            # ניקוי ערכים ריקים
            clean_data = {k: v for k, v in record_data.items() if v != '' and v is not None}
            
            # העלאה
            created_record = self.table.create(clean_data)
            return created_record
            
        except Exception as e:
            # ניתוח שגיאות נפוצות
            error_msg = str(e)
            if "Invalid API key" in error_msg:
                raise Exception("API Key לא תקין. אנא בדוק שהמפתח נכון ופעיל.")
            elif "Base not found" in error_msg:
                raise Exception("Base ID לא נמצא. אנא בדוק שה-Base ID נכון.")
            elif "Table not found" in error_msg:
                raise Exception("הטבלה לא נמצאה. אנא בדוק ש-Table ID נכון.")
            elif "field" in error_msg.lower() and "does not exist" in error_msg.lower():
                raise Exception("שדה לא קיים בטבלה. אנא בדוק את מבנה הטבלה באייר טייבל.")
            else:
                raise Exception(f"שגיאה בהעלאה לאייר טייבל: {error_msg}")
    
    def get_records(self, max_records: int = 100) -> List[Dict]:
        """קבלת רשומות מהטבלה"""
        try:
            if not self.table:
                raise Exception("חיבור לאייר טייבל לא הוקם")
            
            records = self.table.all(max_records=max_records)
            return records
            
        except Exception as e:
            raise Exception(f"שגיאה בקריאת רשומות: {str(e)}")
    
    def update_record(self, record_id: str, record_data: Dict[str, Any]) -> Dict:
        """עדכון רשומה באייר טייבל"""
        try:
            if not self.table:
                raise Exception("חיבור לאייר טייבל לא הוקם")
            
            # ניקוי ערכים ריקים
            clean_data = {k: v for k, v in record_data.items() if v != '' and v is not None}
            
            updated_record = self.table.update(record_id, clean_data)
            return updated_record
            
        except Exception as e:
            raise Exception(f"שגיאה בעדכון רשומה: {str(e)}")
    
    def delete_record(self, record_id: str) -> bool:
        """מחיקת רשומה מאייר טייבל"""
        try:
            if not self.table:
                raise Exception("חיבור לאייר טייבל לא הוקם")
            
            self.table.delete(record_id)
            return True
            
        except Exception as e:
            raise Exception(f"שגיאה במחיקת רשומה: {str(e)}")
    
    def validate_table_structure(self, required_fields: List[str]) -> Dict[str, bool]:
        """בדיקת מבנה הטבלה"""
        try:
            if not self.table:
                return {field: False for field in required_fields}
            
            # ניסיון לקבל רשומה אחת כדי לראות את השדות
            records = self.table.all(max_records=1)
            
            if not records:
                # אם אין רשומות, לא ניתן לבדוק את המבנה
                return {field: None for field in required_fields}
            
            available_fields = set(records[0]['fields'].keys())
            
            return {field: field in available_fields for field in required_fields}
            
        except Exception as e:
            print(f"שגיאה בבדיקת מבנה הטבלה: {e}")
            return {field: False for field in required_fields}

class AirtableManager:
    """מנהל אייר טייבל עם ניהול הגדרות"""
    
    def __init__(self, settings_manager):
        self.settings = settings_manager
        self.client = None
        self._update_client()
    
    def _update_client(self):
        """עדכון הלקוח בהתאם להגדרות"""
        api_key = self.settings.get("airtable.api_key", "")
        base_id = self.settings.get("airtable.base_id", "")
        table_id = self.settings.get("airtable.table_id", "tblC0hR3gZFXxstbM")
        
        if api_key and base_id:
            try:
                self.client = AirtableClient(api_key, base_id, table_id)
            except Exception as e:
                print(f"שגיאה ביצירת לקוח אייר טייבל: {e}")
                self.client = None
        else:
            self.client = None
    
    def is_configured(self) -> bool:
        """בדיקה אם אייר טייבל מוגדר"""
        return self.client is not None
    
    def test_connection(self) -> bool:
        """בדיקת חיבור"""
        if not self.client:
            return False
        return self.client.test_connection()
    
    def upload_results(self, results: List[Dict], file_name: str = "") -> Dict:
        """העלאת תוצאות לאייר טייבל"""
        if not self.client:
            raise Exception("אייר טייבל לא מוגדר. אנא הגדר את הפרטים תחילה.")
        
        from .data_processor import DataProcessor
        processor = DataProcessor()
        
        # יצירת פורמט אייר טייבל
        airtable_data = processor.create_airtable_format(results, file_name)
        
        # העלאה
        return self.client.upload_record(airtable_data)
    
    def refresh_settings(self):
        """רענון ההגדרות"""
        self._update_client()
