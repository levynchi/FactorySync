"""
FactorySync - מערכת לניתוח ועיבוד קבצי אופטיטקס
"""

__version__ = "2.0.0"
__author__ = "FactorySync Team"
__description__ = "מערכת מקצועית לניתוח ועיבוד קבצי אופטיטקס וייצואם לאקסל"

# יבוא הרכיבים הראשיים
from .config.settings import SettingsManager
from .core.file_analyzer import OptitexFileAnalyzer
from .core.data_processor import DataProcessor

__all__ = [
    'SettingsManager',
    'OptitexFileAnalyzer', 
    'DataProcessor'
]
