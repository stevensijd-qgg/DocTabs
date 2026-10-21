@echo off
rem Build DocTabs.exe with the .NET Framework 4.x compiler that ships with Windows
cd /d "%~dp0"
"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /codepage:65001 /target:winexe /optimize+ /out:DocTabs.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll DocTabs.cs
