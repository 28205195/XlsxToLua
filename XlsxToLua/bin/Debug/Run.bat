@echo off
XlsxToLua.exe TestExcel ExportLua -noClient -noLang -columnInfo
set errorLevel = %errorlevel%
if errorLevel == 0 (
	@echo 导出成功
) else (
	@echo 导出失败
)
pause