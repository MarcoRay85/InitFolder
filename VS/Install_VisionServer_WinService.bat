D:
CD D:\Tools\IntactVisionServer
REM der hartkodierte Service-Name ist "IntactVisionService"
%SYSTEMROOT%\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe /unattended IntactVisionServer.WinService.exe
REM sc config IntactVisionService depend= LanmanServer
sc config IntactVisionService depend= tcpip/dhcp/dnscache
sc config IntactVisionService start= delayed-auto
sc failure IntactVisionService actions= restart/30000/restart/30000/restart/30000 reset= 86400

PAUSE
