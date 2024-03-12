'----------------------------------------------------------------對應網站
kill4399 = "http://icsvpn.tcc.net.tw/ICSClientMain.aspx"
kill5001 = "https://icsvpn.tcc.net.tw/ICSClientMain.aspx"
kill5002 = "http://icsvpn.tcc.net.tw/icsinfologin.aspx"
killicsinfologin = "https://icsvpn.tcc.net.tw/icsinfologin.aspx"
'---------------------------------------------------------------指令開始執行

set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then
end if
while(true)
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4399, 1) then wnd.Quit()
if InStr(1, wnd.LocationURL, kill5001, 1) then wnd.Quit()
if InStr(1, wnd.LocationURL, kill5002, 1) then wnd.Quit()
if InStr(1, wnd.LocationURL, killicsinfologin, 1) then wnd.Quit()
next
Wscript.Sleep 50
wend
