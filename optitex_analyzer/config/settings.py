"""
ניהול הגדרות התוכנה
"""

import json
import os
from typing import Dict, Any

class SettingsManager:
    """מנהל הגדרות התוכנה"""
    
    def __init__(self, config_file: str = "config.json"):
        print("🔧 מאתחל מנהל הגדרות...")
        self.config_file = config_file
        print(f"📁 קובץ הגדרות: {self.config_file}")
        
        self.default_config = {
            "app": {
                "window_size": "800x600",
                "auto_load_products": True,
                "products_file": "קובץ מוצרים.xlsx"
            },
            "git": {
                "auto_sync_enabled": False,
                "repo_url": "",
                "branch": "main",
                "last_sync": ""
            },
            "website": {
                "local_url": "http://127.0.0.1:8000",
                "prod_url": "https://arye-textil.co.il",
                "api_token": ""
            }
        }
        print("✅ הגדרות ברירת מחדל נקבעו")
        
        print("📄 טוען הגדרות מהקובץ...")
        self.config = self.load_config()
        print("✅ הגדרות נטענו בהצלחה")
    
    def load_config(self) -> Dict[str, Any]:
        """טעינת הגדרות מהקובץ"""
        try:
            print(f"🔍 בודק אם קובץ הגדרות קיים: {self.config_file}")
            if os.path.exists(self.config_file):
                print("✅ קובץ הגדרות קיים, קורא...")
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    print(f"📝 תוכן הקובץ (תווים): {len(content)}")
                    if content:  # בדיקה שהקובץ לא ריק
                        loaded_config = json.loads(content)
                        print("✅ הגדרות נקראו מהקובץ בהצלחה")
                        # מזיגה עם הגדרות ברירת מחדל
                        merged_config = self._merge_configs(self.default_config, loaded_config)
                        print("✅ הגדרות מוזגו עם ברירת מחדל")
                        return merged_config
                    else:
                        print("⚠️  קובץ הגדרות ריק")
            else:
                print("⚠️  קובץ הגדרות לא קיים")
            
            # אם הקובץ לא קיים או ריק
            print("💾 שומר הגדרות ברירת מחדל...")
            self.save_config(self.default_config)
            print("✅ הגדרות ברירת מחדל נשמרו")
            return self.default_config.copy()
            
        except Exception as e:
            print(f"שגיאה בטעינת הגדרות: {e}")
            return self.default_config.copy()
    
    def save_config(self, config: Dict[str, Any] = None) -> bool:
        """שמירת הגדרות לקובץ"""
        if config is None:
            config = self.config
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"שגיאה בשמירת הגדרות: {e}")
            return False
    
    def get(self, key_path: str, default=None):
        """קבלת ערך לפי נתיב (למשל: 'app.window_size')"""
        keys = key_path.split('.')
        value = self.config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any) -> bool:
        """הגדרת ערך לפי נתיב"""
        keys = key_path.split('.')
        config_ref = self.config
        
        try:
            # יצירת הנתיב אם לא קיים
            for key in keys[:-1]:
                if key not in config_ref:
                    config_ref[key] = {}
                config_ref = config_ref[key]
            
            # הגדרת הערך
            config_ref[keys[-1]] = value
            return self.save_config()
        except Exception as e:
            print(f"שגיאה בהגדרת ערך: {e}")
            return False
    
    def _merge_configs(self, default: Dict, loaded: Dict) -> Dict:
        """מזיגת הגדרות עם שמירה על מבנה ברירת המחדל"""
        result = default.copy()
        
        for key, value in loaded.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_configs(result[key], value)
            else:
                result[key] = value
        
        return result

# אובייקט גלובלי לשימוש בכל התוכנה  
settings = SettingsManager()
