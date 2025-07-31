#!/usr/bin/env python3
"""
סקריפט להצגת כל הטבלאות הזמינות באירטייבל
"""

import json
from pyairtable import Api

def list_tables():
    """מציג את כל הטבלאות הזמינות"""
    
    # טעינת הגדרות
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    api_key = config.get("airtable", {}).get("api_key", "")
    base_id = config.get("airtable", {}).get("base_id", "")
    
    print(f"🏠 בסיס: {base_id}")
    print("📋 טבלאות זמינות:\n")
    
    try:
        api = Api(api_key)
        
        # נסה לקבל מידע על הבסיס
        base_info = api.base(base_id)
        tables = base_info.schema()
        
        if tables and 'tables' in tables:
            for i, table in enumerate(tables['tables'], 1):
                table_id = table.get('id', 'לא ידוע')
                table_name = table.get('name', 'ללא שם')
                
                print(f"{i}. שם: {table_name}")
                print(f"   ID: {table_id}")
                
                # הצגת שדות זמינים
                if 'fields' in table:
                    field_names = [field.get('name', '') for field in table['fields'][:5]]
                    print(f"   שדות: {', '.join(field_names)}...")
                
                print()
        else:
            print("❌ לא הצלחתי לקבל רשימת טבלאות")
            
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        print("\n🔍 ננסה דרך אחרת...")
        
        # ניסיון ישיר לקבל טבלאות
        try:
            # נבדוק כמה table IDs נפוצים
            common_patterns = [
                "tbl" + "0" * 14,  # ברירת מחדל
                "tblC0hR3gZFXxstbM",  # הנוכחי
            ]
            
            print("🔍 בודק טבלאות אפשריות...")
            
            # אפשרות: ננסה לקרוא ישירות מה-base ללא table ID ספציפי
            import requests
            
            headers = {
                'Authorization': f'Bearer {api_key}',
            }
            
            # קריאה למידע על הבסיס
            url = f'https://api.airtable.com/v0/meta/bases/{base_id}/tables'
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                if 'tables' in data:
                    print("✅ נמצאו טבלאות:")
                    for i, table in enumerate(data['tables'], 1):
                        print(f"{i}. {table.get('name', 'ללא שם')} - ID: {table.get('id', 'לא ידוע')}")
                else:
                    print("❌ לא נמצאו טבלאות בתגובה")
            else:
                print(f"❌ שגיאה בקריאת מידע: {response.status_code}")
                print(f"   תגובה: {response.text}")
                
        except Exception as e2:
            print(f"❌ שגיאה גם בניסיון השני: {e2}")

if __name__ == "__main__":
    print("=" * 50)
    print("📋 רשימת טבלאות באירטייבל")
    print("=" * 50)
    list_tables()
