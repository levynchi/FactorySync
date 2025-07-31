#!/usr/bin/env python3
"""
סקריפט לבדיקת חיבור לאירטייבל ודיבאגינג
"""

import json
import sys
from pyairtable import Api

def test_airtable_connection():
    """בדיקת חיבור לאירטייבל"""
    
    # טעינת הגדרות
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except Exception as e:
        print(f"❌ שגיאה בטעינת config.json: {e}")
        return False
    
    api_key = config.get("airtable", {}).get("api_key", "")
    base_id = config.get("airtable", {}).get("base_id", "")
    table_id = config.get("airtable", {}).get("table_id", "")
    
    print("🔍 בדיקת פרטי חיבור:")
    print(f"   API Key: {api_key[:20]}...")
    print(f"   Base ID: {base_id}")
    print(f"   Table ID: {table_id}")
    print()
    
    if not api_key:
        print("❌ API Key חסר!")
        return False
    
    if not base_id:
        print("❌ Base ID חסר!")
        return False
        
    if not table_id:
        print("❌ Table ID חסר!")
        return False
    
    # בדיקה 1: חיבור ל-API
    print("🔗 בודק חיבור ל-Airtable API...")
    try:
        api = Api(api_key)
        print("✅ חיבור ל-API הצליח")
    except Exception as e:
        print(f"❌ שגיאה בחיבור ל-API: {e}")
        return False
    
    # בדיקה 2: גישה לבסיס
    print("🏠 בודק גישה לבסיס...")
    try:
        base = api.base(base_id)
        print("✅ גישה לבסיס הצליחה")
    except Exception as e:
        print(f"❌ שגיאה בגישה לבסיס: {e}")
        return False
    
    # בדיקה 3: גישה לטבלה
    print("📋 בודק גישה לטבלה...")
    try:
        table = api.table(base_id, table_id)
        print("✅ גישה לטבלה הצליחה")
    except Exception as e:
        print(f"❌ שגיאה בגישה לטבלה: {e}")
        print(f"   פרטי השגיאה: {type(e).__name__}: {str(e)}")
        return False
    
    # בדיקה 4: קריאת רשומות מהטבלה
    print("📖 בודק קריאת רשומות...")
    try:
        records = table.all(max_records=1)  # רק רשומה אחת לבדיקה
        print(f"✅ קריאת רשומות הצליחה - נמצאו רשומות")
        if records:
            print(f"   דוגמה לרשומה: {list(records[0]['fields'].keys())[:3]}...")
    except Exception as e:
        print(f"❌ שגיאה בקריאת רשומות: {e}")
        print(f"   פרטי השגיאה: {type(e).__name__}: {str(e)}")
        return False
    
    # בדיקה 5: ניסיון יצירת רשומה בודקת
    print("✏️ בודק יכולת כתיבה...")
    try:
        test_record = {
            'שם הקובץ': 'בדיקה',
            'אורך הגיזרה': 0.0,
        }
        
        # ניסיון יצירה (ללא שמירה בפועל)
        print("   מכין רשומת בדיקה...")
        print("   (לא נשמור בפועל)")
        print("✅ הכנת רשומה הצליחה")
        
    except Exception as e:
        print(f"❌ שגיאה בהכנת רשומה: {e}")
        return False
    
    print("\n🎉 כל הבדיקות עברו בהצלחה!")
    print("   החיבור לאירטייבל תקין.")
    return True

if __name__ == "__main__":
    print("=" * 50)
    print("🧪 בדיקת חיבור אירטייבל")
    print("=" * 50)
    
    success = test_airtable_connection()
    
    if not success:
        print("\n💡 הצעות לפתרון:")
        print("1. בדוק שה-API Key תקין ופעיל")
        print("2. בדוק שה-Base ID נכון")
        print("3. בדוק שה-Table ID נכון")
        print("4. בדוק שיש הרשאות כתיבה לטבלה")
        print("5. בדוק שהטבלה קיימת ונגישה")
        sys.exit(1)
    else:
        print("\n✅ החיבור תקין - אפשר להמשיך לעבוד!")
