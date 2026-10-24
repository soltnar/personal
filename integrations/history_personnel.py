"""Signed, aggregate-only bridge from saved Order History reports to Personnel."""
import hashlib
import hmac
import os
import time
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from urllib.parse import urlparse, parse_qs


def handle_personnel(handler, build_payload):
    parsed = urlparse(handler.path)
    if parsed.path != '/api/personnel-delivery':
        return False
    timestamp = handler.headers.get('X-Personnel-Time', '')
    signature = handler.headers.get('X-Personnel-Signature', '')
    secret = os.environ.get('SABY_APP_SECRET', '')
    try:
        valid_time = abs(time.time() - int(timestamp)) <= 60
    except ValueError:
        valid_time = False
    expected = hmac.new(secret.encode(), ('personnel-delivery\n' + timestamp + '\n' + handler.path).encode(), hashlib.sha256).hexdigest()
    if not secret or not valid_time or not hmac.compare_digest(expected, signature):
        handler.send_json({'error': 'Unauthorized'}, status=401)
        return True
    try:
        qs = parse_qs(parsed.query)
        start, end = date.fromisoformat(qs['from'][0]), date.fromisoformat(qs['to'][0])
        if not 0 <= (end-start).days <= 61:
            raise ValueError('Invalid period')
        from report_store import reports_between
        reports = reports_between(start.isoformat(), end.isoformat())
        present = {report['day'] for report in reports}
        missing = [(start + timedelta(days=i)).isoformat() for i in range((end-start).days+1)
                   if (start + timedelta(days=i)).isoformat() not in present]
        if missing:
            handler.send_json({'error': 'Нет сохранённых дней истории заказов', 'missingDates': missing}, status=409)
            return True
        metrics = []
        for report in reports:
            payload = build_payload(Path(report['filename']))
            # Use the same restaurant/channel split and amounts as the History UI.
            grouped = defaultdict(list)
            for order in payload['orders']:
                restaurant = str(order.get('restaurant') or '')
                if restaurant.endswith(' — Доставка'):
                    grouped[restaurant].append(order)
            for restaurant, orders in grouped.items():
                amounts = [order.get('amount') for order in orders]
                metrics.append({'dateIso': report['day'], 'restaurant': restaurant,
                                'orders': len(orders),
                                'revenue': float(sum((Decimal(str(v)) for v in amounts), Decimal(0))) if all(v is not None for v in amounts) else None})
        handler.send_json({'metrics': metrics, 'completeDates': sorted(present), 'source': 'order-history'})
    except (ValueError, KeyError, IndexError):
        handler.send_json({'error': 'Invalid period'}, status=400)
    except Exception:
        handler.send_json({'error': 'Order history unavailable'}, status=503)
    return True
