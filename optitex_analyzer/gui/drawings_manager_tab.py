"""
Shim module: keep legacy import path working while the implementation lives in gui/drawings_manager/.

Do not add new code here; import the mixin from the feature package.
"""

from .drawings_manager import DrawingsManagerTabMixin

__all__ = ["DrawingsManagerTabMixin"]
