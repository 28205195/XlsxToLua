@echo off
XlsxToLua.exe TestExcel ExportLua -noClient -noLang -columnInfo
set errorLevel = %errorlevel%
if errorLevel == 0 (
	@echo �����ɹ�
) else (
	@echo ����ʧ��
)
pause