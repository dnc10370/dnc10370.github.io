


kill4301 = "https://scms.tcc.net.tw/openchannel/IE/countdown.htm"

set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then

end if
while(true)
for each wnd in wnds

if InStr(1, wnd.LocationURL, kill4301, 1) then wnd.Quit()
Wscript.Sleep 50
next
wend

