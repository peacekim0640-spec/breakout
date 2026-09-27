"""시장 일지 블로그 생성기.

content/posts/*.md (앞머리에 title, date, category, summary) 를 읽어
정적 HTML 사이트를 만든다. GitHub Pages 는 만들어진 파일을 그대로 보여준다.

사용: python3 build.py [출력 폴더, 기본 site]
"""
import html
import re
import shutil
import sys
from datetime import date
from pathlib import Path

import markdown

ROOT = Path(__file__).parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'site'
SITE_NAME = '시장 일지'
SITE_TAGLINE = '구글 파이낸스 데이터로 기록하는 증시·매크로·종목 이슈'
WEEKDAYS = '월화수목금토일'


def parse_post(path):
    text = path.read_text(encoding='utf-8')
    m = re.match(r'^---\n(.*?)\n---\n(.*)$', text, re.S)
    if not m:
        raise ValueError(f'{path.name}: 맨 위에 --- 로 감싼 앞머리가 없습니다')
    meta = {}
    for line in m.group(1).splitlines():
        key, _, value = line.partition(':')
        meta[key.strip()] = value.strip()
    for key in ('title', 'date', 'summary'):
        if not meta.get(key):
            raise ValueError(f'{path.name}: 앞머리에 {key} 가 없습니다')
    body = markdown.markdown(m.group(2), extensions=['tables'])
    d = date.fromisoformat(meta['date'])
    return {
        **meta,
        'slug': path.stem,
        'date_obj': d,
        'date_label': f'{d.year}년 {d.month}월 {d.day}일 ({WEEKDAYS[d.weekday()]})',
        'body': body,
    }


def page(title, body, depth=0):
    prefix = '../' * depth
    return f'''<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Gowun+Batang:wght@400;700&family=IBM+Plex+Sans+KR:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<link rel="stylesheet" href="{prefix}assets/style.css">
</head>
<body>
<div class="wrap">
  <header class="site">
    <a class="brand" href="{prefix}index.html">{SITE_NAME}<small>{SITE_TAGLINE}</small></a>
    <nav aria-label="사이트 메뉴">
      <a href="{prefix}index.html">최신 글</a>
      <a href="{prefix}about.html">소개</a>
    </nav>
  </header>
{body}
  <footer>{SITE_NAME} · 투자 권유가 아닌 기록용 글입니다. 수치는 표시한 기준일 값이며 지연·오류가 있을 수 있습니다.</footer>
</div>
</body>
</html>
'''


def post_article(p):
    return f'''  <article class="post">
    <div class="meta"><span class="cat">{html.escape(p.get('category', '시장'))}</span><time datetime="{p['date']}">{p['date_label']}</time></div>
    <h1>{html.escape(p['title'])}</h1>
    <p class="lede">{html.escape(p['summary'])}</p>
    <div class="body">{p['body']}</div>
  </article>'''


def post_list(posts, depth=0):
    prefix = '../' * depth
    items = '\n'.join(
        f'      <li><time datetime="{p["date"]}">{p["date"].replace("-", ".")}</time>'
        f'<a href="{prefix}posts/{p["slug"]}.html">{html.escape(p["title"])}</a></li>'
        for p in posts
    )
    return f'''    <section>
      <h3>지난 글</h3>
      <ul class="posts">
{items}
      </ul>
    </section>'''


ABOUT = '''  <article class="post">
    <h1>이 블로그는</h1>
    <div class="body">
      <p>매일 장 마감 뒤 구글 파이낸스에서 지수·환율·VIX·관심종목 시세를 모으고, 한 주의 매크로 이슈와 종목 이슈를 정리해 기록합니다.</p>
      <ul>
        <li><b>매일</b> 구글 시트가 시세를 모으고 날짜별로 저장합니다.</li>
        <li><b>매주</b> 모은 데이터와 뉴스를 바탕으로 글 초안을 씁니다.</li>
        <li><b>사람이 읽고 승인</b>하면 이 블로그에 올라갑니다.</li>
      </ul>
      <p>투자 권유가 아닌 기록입니다. 매매 판단과 책임은 읽는 분에게 있습니다.</p>
    </div>
  </article>'''


def main():
    posts = sorted((parse_post(f) for f in (ROOT / 'content/posts').glob('*.md')),
                   key=lambda p: p['date_obj'], reverse=True)
    if OUT.exists():
        shutil.rmtree(OUT)
    (OUT / 'posts').mkdir(parents=True)
    shutil.copytree(ROOT / 'assets', OUT / 'assets')
    (OUT / '.nojekyll').write_text('')  # GitHub Pages 가 Jekyll 로 다시 가공하지 않게

    for i, p in enumerate(posts):
        others = posts[:i] + posts[i + 1:]
        aside = f'  <aside>\n{post_list(others, 1)}\n  </aside>' if others else ''
        body = f'  <div class="layout">\n{post_article(p)}\n{aside}\n  </div>'
        (OUT / 'posts' / f'{p["slug"]}.html').write_text(page(f'{p["title"]} · {SITE_NAME}', body, 1), encoding='utf-8')

    if posts:
        latest = posts[0]
        more = f'<p class="more"><a href="posts/{latest["slug"]}.html">이 글 따로 보기</a></p>'
        aside = f'  <aside>\n{post_list(posts[1:])}\n  </aside>' if len(posts) > 1 else ''
        body = f'  <div class="layout">\n{post_article(latest)}\n{aside}\n  </div>\n  {more}'
    else:
        body = '  <p>아직 글이 없습니다.</p>'
    (OUT / 'index.html').write_text(page(SITE_NAME, body), encoding='utf-8')
    (OUT / 'about.html').write_text(page(f'소개 · {SITE_NAME}', f'  <div class="layout">\n{ABOUT}\n  </div>'), encoding='utf-8')
    print(f'{len(posts)}개 글 → {OUT}')


if __name__ == '__main__':
    main()
