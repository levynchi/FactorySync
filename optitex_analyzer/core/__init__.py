"""רכיבי הליבה של המערכת

NOTE:
    נמנע מייבוא מוקדם של DataProcessor כאן בגלל בעיית יבוא מעגלית שהתגלתה
    (ImportError: cannot import name 'DataProcessor').
    כעת מייצאים רק את מנתח הקבצים; את DataProcessor יש לייבא ישירות מ-
    optitex_analyzer.core.data_processor בעת הצורך.
"""

from .file_analyzer import OptitexFileAnalyzer  # יציב ולא גורר תלות אחורה

__all__ = ['OptitexFileAnalyzer']

# פונקציית עזר לטעינה עצלה (אופציונלי – מי שצריך יכול לקרוא):
def load_data_processor_class():  # pragma: no cover - כלי עזר
        from .data_processor import DataProcessor  # ייבוא בזמן ריצה
        return DataProcessor
