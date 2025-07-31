#!/usr/bin/env python3
"""
בדיקה פשוטה של העלאה לאירטייבל
"""

import json
from pyairtable import Api
import time

def simple_test():
    """בדיקה פשוטה של העלאה"""
    
    # טעינת הגדרות
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    api_key = config.get("airtable", {}).get("api_key", "")
    base_id = config.get("airtable", {}).get("base_id", "")
    table_id = config.get("airtable", {}).get("table_id", "")
    
    try:
        api = Api(api_key)
        table = api.table(base_id, table_id)
        
        print("🧪 מנסה להעלות רשומה פשוטה...")
        
        # רשומה מאוד בסיסית
        simple_record = {
            'שם הקובץ': f'בדיקה-{int(time.time())}'
        }
        
        print(f"   מעלה: {simple_record}")
        created = table.create(simple_record)
        
        print(f"✅ הצלחה! ID רשומה: {created['id']}")
        
        # עכשיו ננסה עם דגם אחד (כטקסט רגיל, לא Linked Record)
        print("\n🧪 מנסה עם דגם אחד...")
        
        record_with_model = {
            'שם הקובץ': f'בדיקה-דגם-{int(time.time())}',
            'דגם 1': 'בייסיק 0-3',  # טקסט רגיל
            'כמות דגם 1': 5
        }
        
        print(f"   מעלה: {record_with_model}")
        created2 = table.create(record_with_model)
        
        print(f"✅ הצלחה! ID רשומה: {created2['id']}")
        
        return True
        
    except Exception as e:
        print(f"❌ שגיאה: {e}")
        return False

if __name__ == "__main__":
    print("=" * 50)
    print("🧪 בדיקה פשוטה של העלאה")
    print("=" * 50)
    simple_test()
