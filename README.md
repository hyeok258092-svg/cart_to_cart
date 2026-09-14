# 오픈마켓 견적서 변환기 (웹)

지마켓·옥션(`.xls`)과 11번가(PDF) 견적서를 엑셀 표로 정리하는 정적 웹앱입니다.
모든 처리는 방문자의 **브라우저 안에서만** 이뤄집니다 — 파일도 API 키도 서버로 전송되지 않습니다.

- 파이썬 로직을 [Pyodide](https://pyodide.org)(WebAssembly)로 브라우저에서 실행합니다.
- 품명 정리는 방문자 본인의 Gemini API 키로 이뤄지며(선택), 키는 브라우저에서 구글로 직접 전송됩니다.
- 키를 넣지 않으면 규칙 기반으로만 정리합니다.

## 구성
- `index.html` — UI + Pyodide 부트스트랩
- `converter.py` — 파싱·정리·엑셀 생성 로직 (브라우저에서 fetch 후 실행)
- `.nojekyll` — GitHub Pages 의 Jekyll 처리 비활성화

## 로컬 실행
```bash
python -m http.server 8123
# http://127.0.0.1:8123 접속
```

## 배포
GitHub Pages: Settings → Pages → Deploy from a branch → `main` / `/ (root)`.
