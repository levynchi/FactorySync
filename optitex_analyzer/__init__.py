"""FactorySync - מערכת לניתוח ועיבוד קבצי אופטיטקס

חשוב: הוסרו ייבואיים גלובליים (Especially DataProcessor) כדי למנוע לולאת יבוא.
ייבוא עמוס ב-__init__ מפעיל side-effects (כמו יצירת settings) לפני שכל המחלקות זמינות.
מעתה יש לבצע:
    from optitex_analyzer.config.settings import SettingsManager
    from optitex_analyzer.core.file_analyzer import OptitexFileAnalyzer
    from optitex_analyzer.core.data_processor import DataProcessor
ישירות בקוד המשתמש (כמו בקובץ main.py).
"""

__version__ = "2.0.1"
__author__ = "FactorySync Team"
__description__ = "מערכת מקצועית לניתוח ועיבוד קבצי אופטיטקס וייצואם לאקסל"

__all__ = [
    '__version__',
    '__author__',
    '__description__'
]
