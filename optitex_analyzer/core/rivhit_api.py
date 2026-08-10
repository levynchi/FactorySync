"""קליינט ל-Rivhit Online API - הכנסת פריטים ישירות לריווחית (פריטים ומלאי).

REST JSON מול https://api.rivhit.co.il/online/RivhitOnlineAPI.svc
אימות: api_token בגוף כל בקשה (מהגדרות ריווחית אונליין -> API).
כל תשובה מכילה error_code / client_message / data.
משתמש ב-urllib מהספרייה הסטנדרטית (ללא תלות חיצונית).
"""

import json
import urllib.error
import urllib.request

BASE_URL = 'https://api.rivhit.co.il/online/RivhitOnlineAPI.svc'
DEFAULT_TIMEOUT = 30  # seconds


class RivhitApiError(Exception):
	"""שגיאה מול ה-API של ריווחית (כולל error_code מהשרת)."""

	def __init__(self, message, error_code=None):
		super().__init__(message)
		self.error_code = error_code


class RivhitOnlineClient:
	"""קליינט קטן מול Rivhit Online API."""

	def __init__(self, api_token: str, timeout: int = DEFAULT_TIMEOUT):
		self.api_token = (api_token or '').strip()
		self.timeout = timeout
		if not self.api_token:
			raise RivhitApiError('לא הוגדר API TOKEN של ריווחית')

	def _request(self, method: str, payload: dict = None) -> dict:
		"""שולח POST ל-method (למשל 'Item.New') ומחזיר את data מהתשובה."""
		body = dict(payload or {})
		body['api_token'] = self.api_token
		data = json.dumps(body, ensure_ascii=False).encode('utf-8')
		req = urllib.request.Request(
			f'{BASE_URL}/{method}',
			data=data,
			headers={'Content-Type': 'application/json; charset=utf-8'},
		)
		try:
			with urllib.request.urlopen(req, timeout=self.timeout) as resp:
				raw = resp.read().decode('utf-8')
		except urllib.error.HTTPError as e:
			# ריווחית מחזירה שגיאות (למשל טוקן שגוי) גם כ-HTTP 401 עם גוף JSON רגיל
			try:
				detail = e.read().decode('utf-8')
				parsed = json.loads(detail)
				message = parsed.get('client_message') or parsed.get('debug_message') or detail[:300]
				raise RivhitApiError(f'{message} (קוד {parsed.get("error_code", e.code)})',
									 error_code=parsed.get('error_code', e.code))
			except (ValueError, AttributeError):
				raise RivhitApiError(f'שגיאת HTTP {e.code} מריווחית ({method})')
		except urllib.error.URLError as e:
			raise RivhitApiError(f'לא ניתן להתחבר לריווחית אונליין:\n{e.reason}')
		try:
			result = json.loads(raw)
		except ValueError:
			raise RivhitApiError(f'ריווחית החזירה תשובה לא צפויה (לא JSON) ב-{method}')
		error_code = result.get('error_code')
		if error_code not in (0, None):
			message = result.get('client_message') or result.get('debug_message') or f'שגיאה {error_code}'
			raise RivhitApiError(f'{message} (קוד {error_code})', error_code=error_code)
		return result.get('data') or {}

	# ===== פריטים =====
	def item_groups(self) -> list:
		"""רשימת קבוצות הפריטים: [{'item_group_id':..., 'item_group_name':...}, ...]."""
		data = self._request('Item.Groups')
		return data.get('item_group_list') or []

	def item_list(self, item_group_id=None) -> list:
		"""רשימת הפריטים בריווחית (אופציונלית מסוננת לפי קבוצה)."""
		payload = {}
		if item_group_id is not None:
			payload['item_group_id'] = int(item_group_id)
		data = self._request('Item.List', payload)
		return data.get('item_list') or []

	def item_new(self, item_name: str, item_part_num: str = '', barcode: str = '',
				 cost_nis=None, sale_nis=None, item_group_id=None, storage_id=None) -> dict:
		"""יצירת כרטיס פריט חדש. מחזיר את data (כולל item_id שהוקצה)."""
		payload = {'item_name': (item_name or '').strip()}
		if not payload['item_name']:
			raise RivhitApiError('חסר שם פריט')
		if (item_part_num or '').strip():
			payload['item_part_num'] = str(item_part_num).strip()
		if (barcode or '').strip():
			payload['barcode'] = str(barcode).strip()
		if cost_nis not in (None, ''):
			payload['cost_nis'] = float(str(cost_nis).replace(',', ''))
		if sale_nis not in (None, ''):
			payload['sale_nis'] = float(str(sale_nis).replace(',', ''))
		if item_group_id not in (None, ''):
			payload['item_group_id'] = int(item_group_id)
		if storage_id not in (None, ''):
			payload['storage_id'] = int(storage_id)
		return self._request('Item.New', payload)

	def item_update(self, item_id, **fields) -> dict:
		"""עדכון כרטיס פריט קיים לפי item_id. fields כמו ב-item_new."""
		payload = {'item_id': int(item_id)}
		name = (fields.get('item_name') or '').strip()
		if name:
			payload['item_name'] = name
		for key in ('item_part_num', 'barcode'):
			val = (fields.get(key) or '').strip() if isinstance(fields.get(key), str) else fields.get(key)
			if val:
				payload[key] = str(val)
		for key in ('cost_nis', 'sale_nis'):
			if fields.get(key) not in (None, ''):
				payload[key] = float(str(fields[key]).replace(',', ''))
		for key in ('item_group_id', 'storage_id'):
			if fields.get(key) not in (None, ''):
				payload[key] = int(fields[key])
		return self._request('Item.Update', payload)
