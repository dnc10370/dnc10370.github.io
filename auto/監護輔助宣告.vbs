on error resume next
kill4399 = "https://scms.tcc.net.tw/ScmsCommon/chkDeclaration.do?&module="
kill4400 = "https://scms.tcc.net.tw/ScmsCommon/chkDeclaration.do?module="


set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
while(true)
if Weekday(date)=0 Or Weekday(date)=0 then
end if

Wscript.Sleep 700
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4399, 1) then wnd.Quit()
next
Wscript.Sleep 700
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4400, 1) then wnd.Quit()
next
Wscript.Sleep 700
wend