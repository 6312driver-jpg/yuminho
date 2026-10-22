# -*- coding: utf-8 -*-
"""카카오톡 '나에게 보내기' 알림.

최초 1회:  python kakao_notify.py setup   (REST API 키 입력 → 로그인 → 토큰 저장)
테스트:    python kakao_notify.py test
"""
import json
import sys
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOKEN = HERE / 'kakao_token.json'
REDIRECT = 'https://localhost'
LINK = 'https://www.tradingview.com/symbols/BTCUSD/'


def _post(url, data, headers=None):
    req = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode(),
                                 headers=headers or {}, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f'{e.code} {e.read().decode(errors="replace")}') from None


def _load():
    return json.loads(TOKEN.read_text(encoding='utf-8'))


def _save(cfg):
    TOKEN.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


def _client(cfg):
    d = {'client_id': cfg['rest_key']}
    if cfg.get('client_secret'):
        d['client_secret'] = cfg['client_secret']
    return d


def setup():
    key = input('카카오 REST API 키: ').strip()
    secret = input('Client Secret (사용 안 하면 그냥 Enter): ').strip()
    url = ('https://kauth.kakao.com/oauth/authorize?response_type=code&scope=talk_message'
           f'&client_id={key}&redirect_uri={urllib.parse.quote(REDIRECT, safe="")}')
    print('\n브라우저에서 카카오 로그인 후 동의하세요.')
    print('"사이트에 연결할 수 없음" 화면이 나오면 정상입니다. 주소창의 주소 전체를 복사하세요.\n')
    webbrowser.open(url)
    back = input('주소창 주소 붙여넣기: ').strip()
    code = urllib.parse.parse_qs(urllib.parse.urlparse(back).query).get('code', [back])[0]
    cfg = {'rest_key': key, 'client_secret': secret}
    tok = _post('https://kauth.kakao.com/oauth/token',
                {'grant_type': 'authorization_code', 'redirect_uri': REDIRECT, 'code': code, **_client(cfg)})
    cfg.update(access_token=tok['access_token'], refresh_token=tok['refresh_token'])
    _save(cfg)
    send('비트코인 멱법칙 알림이 연결되었습니다.')
    print('완료: 카카오톡 "나와의 채팅"을 확인하세요.')


def send(text):
    cfg = _load()
    # 액세스 토큰은 수 시간이면 만료되므로 매번 리프레시 토큰으로 새로 받는다
    tok = _post('https://kauth.kakao.com/oauth/token',
                {'grant_type': 'refresh_token', 'refresh_token': cfg['refresh_token'], **_client(cfg)})
    cfg['access_token'] = tok['access_token']
    if 'refresh_token' in tok:  # 만료가 가까우면 새 리프레시 토큰이 함께 온다
        cfg['refresh_token'] = tok['refresh_token']
    _save(cfg)
    tmpl = {'object_type': 'text', 'text': text[:200],
            'link': {'web_url': LINK, 'mobile_web_url': LINK}, 'button_title': '차트 보기'}
    res = _post('https://kapi.kakao.com/v2/api/talk/memo/default/send',
                {'template_object': json.dumps(tmpl, ensure_ascii=False)},
                {'Authorization': f'Bearer {cfg["access_token"]}'})
    if res.get('result_code') != 0:
        raise RuntimeError(f'카카오톡 전송 실패: {res}')


def configured():
    return TOKEN.exists()


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'setup'
    if cmd == 'setup':
        setup()
    elif cmd == 'test':
        send('카카오톡 알림 테스트')
        print('전송 완료')
    input('\nEnter 를 누르면 닫힙니다.')
