# -*- coding: utf-8 -*-
"""Smoke test for the License bridge methods (offline activation path)."""
import os
import sys
import json
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp()  # redirect session file away from the real one
HERE = os.path.dirname(os.path.abspath(__file__))
BRIDGE = os.path.normpath(os.path.join(HERE, '..', 'Tools', 'py-excel-bridge.py'))

import importlib.util
spec = importlib.util.spec_from_file_location('bridge_under_test', BRIDGE)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

# craft a key with the same algorithm as HostLicense.vb
payload = 'v1|tester@qq.com|pro_yearly|1|4102444800|abc123'
key = 'SPK1-' + m._lic_b64url_encode(payload.encode('utf-8')) + '.' + m._lic_sign(payload)
print('KEY:', key[:40] + '...')

r1 = json.loads(m.m_license_activate_key([json.dumps({'key': key})]))
print('activate valid   :', r1.get('success'), r1.get('code', ''), r1.get('entitlement', {}).get('plan'))

r2 = json.loads(m.m_license_get_local_activation([]))
print('local activation :', r2.get('active'), r2.get('login'))

r3 = json.loads(m.m_license_is_subscription_active([]))
print('subscription gate:', r3.get('active'), r3.get('source'))

bad = key[:-2] + 'xx'
r4 = json.loads(m.m_license_activate_key([json.dumps({'key': bad})]))
print('tampered reject  :', (not r4.get('success')), r4.get('code'))

r5 = json.loads(m.m_license_deactivate_for_refund([]))
print('refund proof     :', r5.get('success'), str(r5.get('code', ''))[:12])

r6 = json.loads(m.m_license_activate_key([json.dumps({'key': key})]))
print('reactivate blocked:', (not r6.get('success')), r6.get('code'))

r7 = json.loads(m.m_license_get_local_activation([]))
print('local after deact :', r7.get('active'), r7.get('reason'))

r8 = json.loads(m.m_license_clear_local_activation([]))
print('clear             :', r8.get('success'))

r9 = json.loads(m.m_license_get_entitlement([]))
print('entitlement       :', r9.get('loggedIn') is False)

r10 = json.loads(m.m_license_ping_health([]))
print('ping health       :', r10.get('ok') is False, r10.get('code'))
