set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks

while(true)
for each colTask in colTasks
if instr(colTask.name,"SCMS�@�~�t��(Open channel) -- �������")<>0 then
colTask.close
end if
next
Wscript.Sleep 100


wend