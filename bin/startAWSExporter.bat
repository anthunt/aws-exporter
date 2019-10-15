@echo off
setlocal

SET RUN_DIR=%~dp0
SET ENV_FILE=conf/env.bat
SET CHECK_TMP_FILE=aws_export_check.tmp
SET PROXY_ALIAS=proxycertforawsexporter

cd %RUN_DIR%

:CHECK_ENV

	IF NOT EXIST %ENV_FILE% GOTO SET_ENV
	GOTO RUN_ENV

:RUN_ENV

	echo.run %ENV_FILE%
	call %ENV_FILE%
	GOTO RUN_AWS_EXPORTER


:RUN_AWS_EXPORTER
	 
	echo.Run AWS Exporter Jar
	
	echo.RunScript "%JRE_HOME%\bin\java.exe" -cp ./*;lib/*; %ENV_TRUST_STORE% %ENV_TRUST_STORE_PASS% anthunt.aws.exporter.AWSExportStarter
	echo.
	
	"%JRE_HOME%\bin\java.exe" -cp ./*;lib/*; %ENV_TRUST_STORE% %ENV_TRUST_STORE_PASS% anthunt.aws.exporter.AWSExportStarter

	GOTO RUN_EXIST


:JRE_HOME_FAIL
	echo.
	echo.Invalid JRE home path. try again.
	GOTO SET_JRE_HOME

:MAKE_CONF
	MKDIR conf

:SET_ENV
	echo.
	echo.You have launched AWS Exporter for the first time.
	echo.Java home setting is required.
	IF NOT EXIST "conf" GOTO :MAKE_CONF

:SET_JRE_HOME

	echo.
	SET /P JRE_HOME="JRE Home Path : "

	IF NOT EXIST "%JRE_HOME%\bin\java.exe" GOTO JRE_HOME_FAIL

	echo.
	echo.The set JRE Home is %JRE_HOME%
	GOTO CHECK_USE_PROXY

:CHECK_USE_PROXY

	echo.
	echo.If you use https protocol behind http proxy
	echo.then need regist proxy certification.
	SET /P USE_PROXY="You need regist certification for proxy?(Y/N) : "

	IF /i "%USE_PROXY%" == "Y" GOTO CHECK_TRUST_STORE
	IF /i "%USE_PROXY%" == "N" GOTO MAKE_ENV
	GOTO CHECK_USE_PROXY

:CHECK_TRUST_STORE

	SET TRUST_STORE_KEYTOOL="%JRE_HOME%\bin\keytool"
	SET TRUST_KEY_STORE="%JRE_HOME%\lib\security\cacerts"

	%TRUST_STORE_KEYTOOL% -list -alias %PROXY_ALIAS% -keystore %TRUST_KEY_STORE% -storepass changeit | find /C "Exception" > %CHECK_TMP_FILE%

	SET /P EXCEPTION_COUNT=<%CHECK_TMP_FILE%
	DEL %CHECK_TMP_FILE%

	IF "%EXCEPTION_COUNT%" == "0" GOTO EXIST_TRUST_STORE
	GOTO SET_CER_FILE

:EXIST_TRUST_STORE

	echo.already registered trust store.
	echo.you don't need registration.
	GOTO MAKE_ENV

:SET_CER_FILE_FAIL

	echo.
	echo.Not exist cer file. try again.
	GOTO SET_CER_FILE

:SET_CER_FILE

	echo.
	echo.Proxy Cert File Path is required.
	echo.
	SET /P CERT_FILE="Proxy .CER File Path : "

	IF NOT EXIST "%CERT_FILE%" GOTO SET_CER_FILE_FAIL
	GOTO REGIST_TRUST_STORE

:REGIST_TRUST_STORE_FAIL

	echo.
	echo.Proxy Certificate registration failed. try again.
	echo. 
	GOTO SET_CER_FILE

:REGIST_TRUST_STORE

	%TRUST_STORE_KEYTOOL% -import -alias %PROXY_ALIAS% -file "%CERT_FILE%" -keystore %TRUST_KEY_STORE% -storepass changeit | find /C "Exception" > %CHECK_TMP_FILE%

	SET /P EXCEPTION_COUNT=<%CHECK_TMP_FILE%
	DEL %CHECK_TMP_FILE%

	IF "%EXCEPTION_COUNT%" == "0" GOTO MAKE_ENV
	GOTO REGIST_TRUST_STORE_FAIL
	

:MAKE_ENV

	echo.Make %ENV_FILE%
	echo.SET JRE_HOME=%JRE_HOME%> %ENV_FILE%
	IF /i "%USE_PROXY%" == "Y" GOTO MAKE_TRUST_STORE_ENV
	GOTO CHECK_ENV

:MAKE_TRUST_STORE_ENV

	echo.Add proxy trust store settings
	echo.SET ENV_TRUST_STORE=-Djavax.net.ssl.trustStore=%TRUST_KEY_STORE%>> %ENV_FILE%
	echo.SET ENV_TRUST_STORE_PASS=-Djava.net.ssl.trustStorePassword=changeit>> %ENV_FILE%
	GOTO CHECK_ENV

:RUN_EXIST

	echo.
	echo.End of AWS Exporter
	