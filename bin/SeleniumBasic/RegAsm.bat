@ECHO OFF
cd /d %~dp0
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "Selenium.dll" /tlb "Selenium.tlb"
PAUSE
CLS