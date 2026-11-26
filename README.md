# FREND Performance Simulator

Streamlit 버전의 Method Comparison 모듈입니다. 합성 검체를 생성하고 Method Comparison 분석 시트와 원자료 시트가 포함된 Excel 파일을 내려받습니다.

## 로컬 실행

Python 3.12 이상에서 다음을 실행합니다.

```text
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

보유한 분석적 성능시험 Excel 양식을 앱 화면에서 선택합니다. 팀이 사용하도록 승인된 빈 양식을 저장소에서 제공하려면 프로젝트 루트에 `FREND_Template.xlsx`를 추가할 수 있습니다. 별도 로컬 경로는 `FREND_TEMPLATE_PATH` 환경 변수로 설정할 수 있습니다.

## 새 비공개 앱 배포

1. 이 디렉터리의 파일을 새 비공개 GitHub 저장소에 올립니다.
2. Streamlit Community Cloud에서 새 앱을 만들고 저장소와 브랜치를 선택합니다. 진입 파일은 `streamlit_app.py`입니다.
3. 앱 공유 설정을 비공개로 두고 허용할 팀원을 추가합니다.

기존 Excel 파일이나 측정 원자료는 이 패키지에 포함되어 있지 않습니다. 다운로드 결과는 생성한 사용자의 세션에 저장됩니다. 별도의 서버 파일 저장 경로나 사용자 간 결과 공유 기능을 사용하지 않습니다.

