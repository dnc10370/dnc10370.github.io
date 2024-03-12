set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks

while(true)
for each colTask in colTasks
if instr(colTask.name,"SCMS作業系統(Open channel) -- 網頁對話")<>0 then
colTask.close
end if
next
Wscript.Sleep 100


wend