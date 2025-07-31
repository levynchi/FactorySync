"""
פונקציות עזר כלליות
"""

from typing import Any, Dict, List, Optional, Union
import re
from datetime import datetime
import json

class DataHelper:
    """עוזר לטיפול בנתונים"""
    
    @staticmethod
    def clean_numeric_value(value: Any) -> float:
        """ניקוי ערך מספרי"""
        if pd.isna(value) or value == '':
            return 0.0
        
        if isinstance(value, (int, float)):
            return float(value)
        
        # ניקוי מחרוזת
        if isinstance(value, str):
            # הסרת רווחים וסימנים מיוחדים
            cleaned = re.sub(r'[^\d.-]', '', str(value))
            try:
                return float(cleaned) if cleaned else 0.0
            except ValueError:
                return 0.0
        
        return 0.0
    
    @staticmethod
    def extract_size_from_text(text: str) -> Optional[str]:
        """חילוץ מידה מטקסט"""
        if not text:
            return None
        
        # דפוסים נפוצים של מידות
        patterns = [
            r'\b(\d+(?:\.\d+)?)\b',  # מספר בסיסי
            r'size\s*[:=]?\s*(\d+(?:\.\d+)?)',  # Size: XX
            r'מידה\s*[:=]?\s*(\d+(?:\.\d+)?)',  # מידה: XX
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(text), re.IGNORECASE)
            if match:
                return match.group(1)
        
        return None
    
    @staticmethod
    def normalize_product_name(name: str) -> str:
        """נרמול שם מוצר"""
        if not name:
            return ""
        
        # הסרת רווחים מיותרים
        normalized = re.sub(r'\s+', ' ', str(name).strip())
        
        # המרה לאותיות קטנות לצורך השוואה
        normalized = normalized.lower()
        
        # הסרת תווים מיוחדים
        normalized = re.sub(r'[^\w\s\u0590-\u05FF]', '', normalized)
        
        return normalized
    
    @staticmethod
    def format_quantity(quantity: float, decimal_places: int = 1) -> str:
        """עיצוב כמות"""
        if quantity == int(quantity):
            return str(int(quantity))
        else:
            return f"{quantity:.{decimal_places}f}"
    
    @staticmethod
    def parse_size_range(size_text: str) -> List[str]:
        """פירוק טווח מידות"""
        if not size_text:
            return []
        
        # דפוס של טווח: 50-60
        range_match = re.match(r'(\d+)-(\d+)', str(size_text))
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            return [str(i) for i in range(start, end + 1)]
        
        # דפוס של רשימה: 50,52,54
        if ',' in str(size_text):
            return [s.strip() for s in str(size_text).split(',') if s.strip().isdigit()]
        
        # מידה יחידה
        if str(size_text).strip().isdigit():
            return [str(size_text).strip()]
        
        return []

class ValidationHelper:
    """עוזר לוולידציה"""
    
    @staticmethod
    def is_valid_email(email: str) -> bool:
        """בדיקת תקינות מייל"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, str(email)))
    
    @staticmethod
    def is_valid_url(url: str) -> bool:
        """בדיקת תקינות URL"""
        pattern = r'^https?://[^\s/$.?#].[^\s]*$'
        return bool(re.match(pattern, str(url)))
    
    @staticmethod
    def validate_airtable_config(config: Dict[str, Any]) -> tuple[bool, str]:
        """בדיקת תקינות הגדרות אייר טייבל"""
        required_fields = ['api_key', 'base_id', 'table_name']
        
        for field in required_fields:
            if field not in config or not config[field]:
                return False, f"חסר שדה: {field}"
        
        # בדיקת פורמט API Key
        if not config['api_key'].startswith('pat'):
            return False, "API Key חייב להתחיל ב-pat"
        
        # בדיקת פורמט Base ID
        if not config['base_id'].startswith('app'):
            return False, "Base ID חייב להתחיל ב-app"
        
        return True, "תקין"

class FormatHelper:
    """עוזר לעיצוב נתונים"""
    
    @staticmethod
    def format_datetime(dt: datetime, format_str: str = "%d/%m/%Y %H:%M") -> str:
        """עיצוב תאריך ושעה"""
        try:
            return dt.strftime(format_str)
        except Exception:
            return str(dt)
    
    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """עיצוב גודל קובץ"""
        try:
            if size_bytes < 1024:
                return f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.1f} KB"
            elif size_bytes < 1024 * 1024 * 1024:
                return f"{size_bytes / (1024 * 1024):.1f} MB"
            else:
                return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
        except Exception:
            return "Unknown"
    
    @staticmethod
    def truncate_text(text: str, max_length: int = 50) -> str:
        """קיצור טקסט"""
        if not text or len(text) <= max_length:
            return text
        
        return text[:max_length-3] + "..."
    
    @staticmethod
    def format_product_display(product_name: str, max_length: int = 30) -> str:
        """עיצוב הצגת שם מוצר"""
        if not product_name:
            return "לא ידוע"
        
        # הסרת רווחים מיותרים
        formatted = re.sub(r'\s+', ' ', product_name.strip())
        
        # קיצור אם יותר מדי ארוך
        if len(formatted) > max_length:
            formatted = formatted[:max_length-3] + "..."
        
        return formatted

class MathHelper:
    """עוזר לחישובים מתמטיים"""
    
    @staticmethod
    def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
        """חלוקה בטוחה"""
        try:
            if denominator == 0:
                return default
            return numerator / denominator
        except Exception:
            return default
    
    @staticmethod
    def round_to_precision(value: float, precision: int = 1) -> float:
        """עיגול לדיוק מסוים"""
        try:
            return round(value, precision)
        except Exception:
            return 0.0
    
    @staticmethod
    def calculate_percentage(part: float, total: float) -> float:
        """חישוב אחוז"""
        return MathHelper.safe_divide(part * 100, total, 0.0)

# יבוא pandas בצורה בטוחה
try:
    import pandas as pd
except ImportError:
    # אם pandas לא זמין, ניצור dummy class
    class pd:
        @staticmethod
        def isna(value):
            return value is None or (isinstance(value, str) and value.strip() == '')

# פונקציות שימושיות נוספות
def safe_json_loads(json_str: str, default: Any = None) -> Any:
    """טעינת JSON בטוחה"""
    try:
        return json.loads(json_str)
    except Exception:
        return default

def safe_json_dumps(obj: Any, default: str = "{}") -> str:
    """המרת JSON בטוחה"""
    try:
        return json.dumps(obj, ensure_ascii=False, indent=2)
    except Exception:
        return default

def get_nested_value(data: Dict[str, Any], key_path: str, default: Any = None) -> Any:
    """קבלת ערך מקושר בלקסיקון"""
    try:
        keys = key_path.split('.')
        value = data
        for key in keys:
            value = value[key]
        return value
    except (KeyError, TypeError):
        return default

def set_nested_value(data: Dict[str, Any], key_path: str, value: Any) -> None:
    """הגדרת ערך מקושר בלקסיקון"""
    try:
        keys = key_path.split('.')
        current = data
        
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        
        current[keys[-1]] = value
    except Exception:
        pass
