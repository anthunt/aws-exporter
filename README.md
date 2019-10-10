# AWS Exporter 사용방법

# 초기 환경 설정
1. startAWSExporter.bat 파일 및 aws-exporter-x.x.x-RELEASE.jar 파일을 원하는 로컬 경로에 복사
2. startAWSExporter.bat 파일 실행
3. 초기 실행 시 환경 설정을 위한 입력 값들을 설정
4. 프로그램 실행에 따른 입력 값 입력 (JRE Home path 및 Proxy 인증서 설정)
5. 접속 번호를 입력하여 주십시요. 메세지 출력이 되면 설정 완료


# 접속 정보 관리
1. 접속 번호를 입력하여 주십시요. 메시지 출력 후 0번 입력 시 신규 접속 정보 생성
2. 접속명 : 접속 정보를 구분하기 위한 구분값 (예. cfota-prod)
3. AccessKey : AWS IAM 계정 AccessKey
4. SecretKey : AWS IAM 계정 SecretKey
5. Proxy 사용 여부 : Proxy 사용하여 인터넷 연결 여부 (Http Proxy 설정만 지원)
6. Proxy Host : Proxy Host 주소
7. Proxy Port : Proxy Port 번호
8. 접속 정보 생성 완료 시 새로 설정된 모든 접속 정보 리스트를 다시 표시해줌.

# 추출 실행
1. 접속 번호를 입력하여 주십시요. 메시지 출력 후 추출을 원하는 접속 정보의 번호를
   입력하면 추출 시작
2. 추출 완료 후 프로그램 종료


* 추출된 파일은 [접속명][Region명] AWSExport-추출시간.xlsx 로 exports 경로 하위에 생성 됨.

# 프로그램 구조
- startAWSExporter.bat : AWS Exporter 프로그램 시작 파일
- aws-export-x.x.x-RELEASE.jar : AWS Exporter JAR 라이브러리
- env.bat : 실행 환경 변수 파일 (startAWSExporter.bat 파일에 의해서 최초 설정 시 자동으로 생성 됨.)
- exports : 추출 실행 결과 파일 저장 폴더(최초 추출 시 자동으로 생성 됨.)
