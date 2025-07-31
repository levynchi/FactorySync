"""
פונקציות עזר לעבודה עם קבצים
"""

import os
import pandas as pd
from typing import Optional, List, Dict, Any
import shutil
from datetime import datetime

class FileHandler:
    """מחלקה לטיפול בפעולות קבצים"""
    
    @staticmethod
    def validate_excel_file(file_path: str) -> tuple[bool, str]:
        """בדיקת תקינות קובץ אקסל"""
        try:
            if not os.path.exists(file_path):
                return False, "הקובץ לא קיים"
            
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                return False, "הקובץ חייב להיות מסוג Excel"
            
            # ניסיון לקרוא את הקובץ
            df = pd.read_excel(file_path, nrows=5)
            if df.empty:
                return False, "הקובץ ריק"
            
            return True, "תקין"
            
        except Exception as e:
            return False, f"שגיאה בקריאת הקובץ: {str(e)}"
    
    @staticmethod
    def get_excel_sheets(file_path: str) -> List[str]:
        """קבלת רשימת גיליונות בקובץ אקסל"""
        try:
            xl_file = pd.ExcelFile(file_path)
            return xl_file.sheet_names
        except Exception:
            return []
    
    @staticmethod
    def backup_file(file_path: str) -> Optional[str]:
        """יצירת גיבוי של קובץ"""
        try:
            if not os.path.exists(file_path):
                return None
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name, file_ext = os.path.splitext(file_path)
            backup_path = f"{file_name}_backup_{timestamp}{file_ext}"
            
            shutil.copy2(file_path, backup_path)
            return backup_path
            
        except Exception:
            return None
    
    @staticmethod
    def clean_temp_files(directory: str, pattern: str = "*temp*") -> int:
        """ניקוי קבצים זמניים"""
        try:
            import glob
            temp_files = glob.glob(os.path.join(directory, pattern))
            count = 0
            
            for file_path in temp_files:
                try:
                    os.remove(file_path)
                    count += 1
                except Exception:
                    continue
            
            return count
            
        except Exception:
            return 0
    
    @staticmethod
    def get_file_info(file_path: str) -> Dict[str, Any]:
        """קבלת מידע על קובץ"""
        try:
            if not os.path.exists(file_path):
                return {}
            
            stat = os.stat(file_path)
            return {
                'size': stat.st_size,
                'modified': datetime.fromtimestamp(stat.st_mtime),
                'created': datetime.fromtimestamp(stat.st_ctime),
                'name': os.path.basename(file_path),
                'extension': os.path.splitext(file_path)[1],
                'directory': os.path.dirname(file_path)
            }
            
        except Exception:
            return {}
    
    @staticmethod
    def ensure_directory(directory: str) -> bool:
        """וידוא שהתיקייה קיימת"""
        try:
            os.makedirs(directory, exist_ok=True)
            return True
        except Exception:
            return False
    
    @staticmethod
    def safe_filename(filename: str) -> str:
        """יצירת שם קובץ בטוח"""
        # הסרת תווים בלתי חוקיים
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # הגבלת אורך
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
        
        return filename
