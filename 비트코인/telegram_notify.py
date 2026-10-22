# -*- coding: utf-8 -*-
"""텔레그램 봇 알림.

최초 1회:  python telegram_notify.py setup   (봇 토큰 입력 → 봇에게 메시지 → chat_id 저장)
테스트:    python telegram_notify.py test
"""
import json
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = HERE / 'telegram_config.json'


def _call(token, method, **params):
    req = urllib.request.Request(f'https://api.telegram.org/bot{token}/{method}',
                                 data=urllib.parse.urlencode(params).encode(), method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            res = json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f'{e.code} {e.read().decode(errors="replace")}') from None
    if not res.get('ok'):
        raise RuntimeError(f'텔레그램 오류: {res}')
    return res['result']


def setup():
    token = input('BotFather 에게 받은 봇 토큰: ').strip()
    me = _call(token, 'getMe')
    print(f'\n텔레그램에서 @{me["username"]} 봇을 열고 /start 를 보내세요. (최대 2분 대기)')
    chat_id = None
    for _ in range(24):
        for u in _call(token, 'getUpdates', timeout=5):
            msg = u.get('message') or u.get('edited_message')
            if msg:
                chat_id = msg['chat']['id']
        if chat_id:
            break
        time.sleep(5)
    if not chat_id:
        print('메시지를 받지 못했습니다. 다시 실행하세요.')
        return
    CONFIG.write_text(json.dumps({'token': token, 'chat_id': chat_id}, indent=2), encoding='utf-8')
    send('비트코인 멱법칙 알림이 연결되었습니다.')
    print('완료: 텔레그램을 확인하세요.')


def send(text):
    cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
    _call(cfg['token'], 'sendMessage', chat_id=cfg['chat_id'], text=text)


def configured():
    return CONFIG.exists()


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'setup'
    if cmd == 'setup':
        setup()
    elif cmd == 'test':
        send('텔레그램 알림 테스트')
        print('전송 완료')
    input('\nEnter 를 누르면 닫힙니다.')
