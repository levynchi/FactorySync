"""קליינט לייצוא מוצרים לאתר הקטלוג הלבן (arye-textile-branding, Django).

משתמש ב-urllib מהספרייה הסטנדרטית (ללא תלות חיצונית).
האימות: טוקן משותף בכותרת X-API-Token, מוגדר באתר דרך WHITE_CATALOG_API_TOKEN.
"""

import base64
import io
import json
import os
import urllib.error
import urllib.request

DEFAULT_TIMEOUT = 60  # seconds — image payloads need more time


def encode_product_image_b64(path, max_side=1400):
	"""Encode a product photo as JPEG base64 for the website import API.

	Returns (b64, 'jpg') or ('', '') if the file is missing or unreadable.
	"""
	rel = (path or '').strip()
	if not rel:
		return '', ''
	full = rel if os.path.isabs(rel) else os.path.join(os.getcwd(), rel)
	if not os.path.exists(full):
		return '', ''
	try:
		from PIL import Image
		img = Image.open(full).convert('RGB')
		w, h = img.size
		longest = max(w, h)
		if longest > max_side:
			scale = max_side / float(longest)
			img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
		buf = io.BytesIO()
		img.save(buf, 'JPEG', quality=88)
		return base64.b64encode(buf.getvalue()).decode('ascii'), 'jpg'
	except Exception:
		return '', ''


class WebsiteExportError(Exception):
	"""שגיאה בתקשורת מול האתר (כולל תשובת שגיאה מהשרת)."""


class WebsiteClient:
	"""קליינט קטן מול ה-API של הקטלוג הלבן."""

	def __init__(self, base_url: str, token: str, timeout: int = DEFAULT_TIMEOUT):
		self.base_url = (base_url or '').strip().rstrip('/')
		self.token = (token or '').strip()
		self.timeout = timeout
		if not self.base_url:
			raise WebsiteExportError('לא הוגדרה כתובת אתר')

	def _request(self, path: str, payload: dict = None) -> dict:
		url = f"{self.base_url}{path}"
		data = None
		# Cloudflare חוסם את ברירת המחדל Python-urllib (error 1010) - חובה UA מותאם
		headers = {'X-API-Token': self.token, 'User-Agent': 'FactorySync/1.0'}
		if payload is not None:
			data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
			headers['Content-Type'] = 'application/json; charset=utf-8'
		req = urllib.request.Request(url, data=data, headers=headers)
		try:
			with urllib.request.urlopen(req, timeout=self.timeout) as resp:
				body = resp.read().decode('utf-8')
		except urllib.error.HTTPError as e:
			try:
				err_body = e.read().decode('utf-8')
				err_json = json.loads(err_body)
				detail = err_json.get('error') or err_body
			except Exception:
				detail = str(e)
			if e.code == 403:
				raise WebsiteExportError(f'האתר דחה את הבקשה (403): טוקן API שגוי או לא מוגדר.\n{detail}')
			raise WebsiteExportError(f'שגיאה מהאתר ({e.code}): {detail}')
		except urllib.error.URLError as e:
			raise WebsiteExportError(f'לא ניתן להתחבר לאתר בכתובת {self.base_url}:\n{e.reason}')
		try:
			return json.loads(body)
		except ValueError:
			raise WebsiteExportError('האתר החזיר תשובה לא צפויה (לא JSON)')

	def fetch_meta(self) -> dict:
		"""שליפת רשימת המוצרים, סוגי הבדים והמידות מהאתר."""
		return self._request('/white-catalog/api/export-meta/')

	def send_variants(self, product_id: int, fabric_type: str, rows: list) -> dict:
		"""שליחת וריאנטים (מידה+ברקוד+מחיר) למוצר קיים באתר.

		rows: רשימת מילונים עם המפתחות size, barcode, unit_price (ואופציונלית fabric_type).
		למוצרים שנוצרו לפי צבעים השורות כוללות גם color ו-color_hex; מוצרים באתר
		שתומכים בצבע יקלטו אותם, ואחרים יתעלמו מהשדות.
		מחזיר את סיכום השרת: created / updated / errors / warnings.
		"""
		payload = {
			'product_id': product_id,
			'fabric_type': (fabric_type or '').strip(),
			'rows': rows,
		}
		return self._request('/white-catalog/api/variants/import/', payload)
