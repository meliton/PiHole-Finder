@ECHO OFF
CLS
echo Finding probable gateway / router IP(s)
echo.
ipconfig |find "Default"
echo.
echo.
echo Finding Pi-Hole in the network...
echo.
arp -a |find "b8-27-eb"
echo.
echo.
pause

