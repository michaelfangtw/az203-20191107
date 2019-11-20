Option Explicit
dim computer,processName,objWMIService
dim logPath,colItems
computer = "." 
processName="chrome1.exe"
logPath="d:\log\checkFTPFun_" & formatdate("yyyymmdd",now) & ".log"
Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & _
           " WHERE Name = '"& processName &"'",,48) 

dim fso ,stdout,objItem
set fso= Wscript.CreateObject("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream (1)
dim count
count=0       

MyLog ""
MyLog "<START>"
MyLog "check process =" & processName

For Each objItem in colItems     
    count=count+1
Next

if count>1 then
    MyLog "check OK"
    MyLog "count=" & count   
else
        MyLog "chjeck Failed"
        MyLog "count=" & count   
end if
MyLog "<END>"   
MyLog ""

dim result
dim url
url="https://support.oneskyapp.com/hc/en-us/article_attachments/202761627/example_1.json"
result=httpGet(url)
MyLog "httpget,url="  & url
MyLog "body=" & result

   
'=====================================================================================================

public function httpGet(url)
    on error resume next
    dim http
    set http=createObject("Microsoft.XMLHTTP")
    http.open "GET",url,false
    http.send
    dim result

    If http.Status = 200 Then
        result= http.responseText            
    End If
    if err then    
        MyLog "url=" & url &",error=" & err.description
    end if
    httpGet= result
end function

Public Function MyLog(Message )
        dim msg
        msg= formatdate("yyyy/mm/dd hh:nn:ss",now) & chr(9) & message
        call WriteLog(logpath, 8, msg )
End Function


Public Function WriteLog(FileStr, mode, Message )
        Dim fso, f
        Set fso = CreateObject("Scripting.fileSystemObject")
		Set f = fso.OpenTextFile(FileStr, CInt(mode), True)   'mode:8 Appending,mode:2 Writing
        f.WriteLine Message
        f.Close
        Set fso = Nothing
End Function


'********************************
'程序:formatdate
'說明:日期格式化
'時間:2002/05/01
'作者:fij
'傳入:string:strformat,string:strdate
'回傳:日期字串
'範例:formatdate("yyyy/mm/dd hh:nn:ss",now)
'說明:strformat支援格式有
'     yyyy/mm/dd hh:nn:ss
'********************************
function formatdate(strformat,strdate)
dim result
dim yyyy,mm,dd,hh,nn,ss,yy
'yyyy/mm/dd hh:nn:ss
'response.write isdate(strdate)
if isdate(strdate) then
    yy=right("0"&year(strdate),2)
    yyyy=right("000"&year(strdate),4)
    mm=right("0"&month(strdate),2)
    dd=right("0"&day(strdate),2)
    hh=right("0"&Hour(strdate),2)
    nn=right("0"&minute(strdate),2)
    ss=right("0"&Second(strdate),2)   
   result=replace(strformat,"yyyy",yyyy)
   result=replace(result,"yy",yy)
   result=replace(result,"mm",mm)
   result=replace(result,"dd",dd)
   result=replace(result,"hh",hh)
   result=replace(result,"nn",nn)
   result=replace(result,"ss",ss)
else
    result=""
end if
formatdate=result
end function
