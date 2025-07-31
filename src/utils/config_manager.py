"""
מודול לניהול הגדרות התוכנה
"""

import json
import os


class ConfigManager:
    """מחלקה לניהול הגדרות התוכנה"""
    
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.config = {}
        self.load_config()
    
    def get_default_config(self):
        """קבלת הגדרות ברירת מחדל"""
        return {
            "airtable": {
                "api_key": "",
                "base_id": "",
                "table_id": "tblC0hR3gZFXxstbM"
            },
            "app": {
                "check_tubular": True,
                "only_positive_quantities": True,
                "auto_find_products_file": True,
                "products_file_name": "קובץ מוצרים.xlsx"
            }
        }
    
    def load_config(self):
        """טעינת הגדרות מקובץ JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                # יצירת קובץ הגדרות ברירת מחדל
                self.config = self.get_default_config()
                self.save_config()
        except Exception as e:
            print(f"שגיאה בטעינת הגדרות: {e}")
            self.config = self.get_default_config()
    
    def save_config(self):
        """שמירת הגדרות לקובץ JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"שגיאה בשמירת הגדרות: {e}")
            return False
    
    def get_airtable_config(self):
        """קבלת הגדרות Airtable"""
        return self.config.get("airtable", {})
    
    def set_airtable_config(self, api_key, base_id, table_id):
        """עדכון הגדרות Airtable"""
        if "airtable" not in self.config:
            self.config["airtable"] = {}
        
        self.config["airtable"]["api_key"] = api_key.strip()
        self.config["airtable"]["base_id"] = base_id.strip()
        self.config["airtable"]["table_id"] = table_id.strip()
        
        return self.save_config()
    
    def get_app_config(self):
        """קבלת הגדרות האפליקציה"""
        return self.config.get("app", {})
    
    def set_app_config(self, **kwargs):
        """עדכון הגדרות האפליקציה"""
        if "app" not in self.config:
            self.config["app"] = {}
        
        for key, value in kwargs.items():
            self.config["app"][key] = value
        
        return self.save_config()
    
    def get_airtable_status(self):
        """קבלת סטטוס חיבור Airtable"""
        airtable_config = self.get_airtable_config()
        api_key = airtable_config.get("api_key", "")
        base_id = airtable_config.get("base_id", "")
        
        if api_key and base_id:
            return {
                'status': 'connected',
                'text': "🟢 מחובר לאייר טייבל",
                'color': "green"
            }
        elif api_key or base_id:
            return {
                'status': 'partial',
                'text': "🟡 הגדרות חלקיות",
                'color': "orange"
            }
        else:
            return {
                'status': 'disconnected',
                'text': "🔴 לא מוגדר",
                'color': "red"
            }
    
    def find_products_file(self):
        """חיפוש אוטומטי של קובץ מוצרים"""
        app_config = self.get_app_config()
        if not app_config.get("auto_find_products_file", True):
            return None
        
        products_file_name = app_config.get("products_file_name", "קובץ מוצרים.xlsx")
        
        if os.path.exists(products_file_name):
            return os.path.abspath(products_file_name)
        
        return None
