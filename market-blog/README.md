# 시장 일지 (블로그)

구글 파이낸스 데이터와 한 주의 매크로·종목 이슈를 정리해 GitHub Pages 에 올리는 블로그입니다.

```
content/posts/YYYY-MM-DD-제목.md   ← 글 (앞머리: title, date, category, summary)
data/snapshots/YYYY-MM-DD.json    ← 그날 구글 파이낸스에서 모은 숫자
assets/style.css                  ← 디자인
build.py                          ← 글을 HTML 사이트로 만든다
```

## 글 올리기

```bash
pip install markdown
python3 build.py          # site/ 에 사이트가 만들어진다
```

`site/` 안의 파일을 `peacekim0640-spec.github.io` 저장소에 올리면 1~2분 뒤 블로그에 반영됩니다.
