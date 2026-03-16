# PPTX → Markdown/DB 변환 에디터

업로드한 PPTX에서 텍스트, 링크, 이미지, 첨부(embeddings/ole)를 추출해
AI 에이전트 분석용 데이터셋으로 변환합니다.

## 기능
- 슬라이드별 텍스트/노트 추출
- 하이퍼링크 URL 추출
- 이미지 파일 추출(`images/`)
- 첨부파일(embeddings/ole) 원본 추출(`attachments/`)
- 결과물 생성:
  - `document.md` (사람이 읽기 좋은 문서)
  - `slides.jsonl` (슬라이드 단위 레코드)
  - `assets.jsonl` (이미지/첨부 메타)
  - `manifest.json` (요약 메타)

## 실행
```bash
cd ~/Desktop/"새 폴더"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

브라우저에서 파일 업로드 후 `변환 실행`을 누르면 ZIP으로 다운로드할 수 있습니다.

## DB 적재 권장
- 벡터DB: `slides.jsonl`의 `text`를 임베딩
- 검색 인덱스: `title + text + links + notes`
- 자산 참조: `assets.jsonl`의 `file_path` 연결
