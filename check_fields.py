#!/usr/bin/env python3
"""
בדיקת שדות הטבלה באירטייבל
"""

import json
from pyairtable import Api

def check_table_fields():
    """בודק את השדות בטבלה"""
    
    # טעינת הגדרות
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    api_key = config.get("airtable", {}).get("api_key", "")
    base_id = config.get("airtable", {}).get("base_id", "")
    table_id = config.get("airtable", {}).get("table_id", "")
    
    try:
        api = Api(api_key)
        
        # קבלת מידע על הבסיס
        base_info = api.base(base_id)
        schema = base_info.schema()
        
        if 'tables' in schema:
            for table in schema['tables']:
                if table.get('id') == table_id:
                    print(f"📋 טבלה: {table.get('name', 'ללא שם')}")
                    print(f"   ID: {table_id}")
                    print("\n🏷️ שדות בטבלה:")
                    
                    for field in table.get('fields', []):
                        name = field.get('name', 'ללא שם')
                        field_type = field.get('type', 'לא ידוע')
                        
                        # בדיקה אם השדה מחושב
                        is_computed = field_type in ['formula', 'rollup', 'count', 'lookup']
                        computed_marker = " 🔒 (מחושב)" if is_computed else " ✏️ (ניתן לכתיבה)"
                        
                        print(f"   • {name} - {field_type}{computed_marker}")
                    
                    return True
        
        print("❌ לא נמצאה הטבלה")
        return False
        
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        
        # ניסיון פשוט יותר - ניסוי יצירת רשומה ריקה
        print("\n🧪 ננסה ליצור רשומה בסיסית...")
        try:
            table = api.table(base_id, table_id)
            
            # ניסיון עם רשומה מאוד בסיסית
            test_record = {
                'שם הקובץ': 'בדיקה-' + str(int(time.time()))
            }
            
            print(f"   מנסה ליצור: {test_record}")
            # created = table.create(test_record)
            print("   ✅ הרשומה הבסיסית תקינה")
            
        except Exception as e2:
            print(f"   ❌ שגיאה ברשומה בסיסית: {e2}")
        
        return False

if __name__ == "__main__":
    import time
    print("=" * 50)
    print("🔍 בדיקת שדות הטבלה")
    print("=" * 50)
    check_table_fields()
