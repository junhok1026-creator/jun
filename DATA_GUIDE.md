# 기관 데이터 수집 가이드

## 기본 원칙
- 기관 주소/연락처: 고용24 기관찾기, 고용복지+센터 공식 페이지에서 수집
- 업무 분류: 각 센터 `센터에서 하는 일`, `부서 및 직원소개`에서 수집
- 지도 링크: 주소 기반으로 구글/네이버/카카오 링크 생성

## 권장 출처
- 고용24 기관찾기: https://m.work24.go.kr/cm/c/d/0190/retrieveInstSrchLst.do?issn=G0000459
- 서울고용복지+센터(전국 링크 허브 포함): https://www.work.go.kr/seoul/ctrIntro/guideMap/guideMap.do
- 인천고용센터 예시: https://www.work.go.kr/incheon/ctrIntro/deptStaffInfo/deptStaffInfoList.do
- 안산지청(지청 업무/연락처 예시): https://www.moel.go.kr/local/ansan/introduce/direction/list.do
- 지방고용노동청·지청 조직/부서(근로개선지도과, 산재예방지도과 등): https://www.moel.go.kr/menu?menuKey=HBvfxuJIlx
- 산재보험(근로복지공단) 연계: https://devkeupyeo.comwel.or.kr/

## 데이터 필드
- id, name, type
- region.sido, region.sigungu
- jurisdiction[]
- address, tel
- services[]
- source

선택 필드(부서 단위 분리):
- departments[]: `{ name, address, tel, services[], source }`
- 예: 지역협력과(외국인업무)가 본청과 별도 위치면 `departments`에 별도 주소로 추가

## 업무 분장 반영 팁
- 고용센터 업무: 실업급여, 국민취업지원, 취업알선, 직업훈련, 기업지원
- 노동지청 업무: 임금체불, 근로감독, 직장내괴롭힘, 산업안전/산재예방, 중대재해 조사
- 산재보상 급여 신청/지급은 `근로복지공단` 연계 항목으로 별도 관리
- 외국인업무(EPS/권익)는 기관 공통으로 일괄 부여하지 말고, 실제 담당 부서에만 `services` 또는 `departments.services`로 입력

## 빠른 운영 방법
1. 기본 전국 데이터 파일 사용: `data/offices.national.full.json` / `data/offices.national.full.js`
2. 웹 화면 실행: `office_locator.html`
3. 수집 갱신: `powershell -ExecutionPolicy Bypass -File scripts/refresh_offices_data.ps1`
4. 부서 예외/보정: `data/office_overrides.json` 수정 후 스크립트 재실행

## 정확도 높이는 수집 방법(권장)
1. 기관 기본정보 자동수집: `work.go.kr/{센터}/main.do`에서 기관명/주소/전화/관할 수집
2. 업무 카테고리 자동분류: 각 센터 `deptStaffInfoList.do`의 담당업무 문구를 키워드 매핑해 `services` 생성
3. 예외기관 수동보정: 노동청/지청, 근로복지공단, 외국인 전담부서 등은 `office_overrides.json`에 직접 반영
4. 저장 전 안전검증: 스크립트가 링크/센터 건수가 최소 기준 미만이면 파일 덮어쓰기를 자동 중단
5. 표본 검증: 매 회차마다 5~10개 기관을 골라 공식 페이지와 주소/전화/업무를 대조 확인

## 전국 데이터 기준
- 수집 기준일: 2026-04-09
- 고용센터/고용복지+센터: `work.go.kr` 전국센터 링크 132개 자동 수집
- 추가 기관: 노동지청/근로복지공단(산재보상) 샘플 4개
- 현재 자동수집 결과: 총 136개, 주소/전화/시도 누락 0건
