<!--#include file="CheckAdmin.asp"-->
<!--#include file="Admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|34,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
Function htmll(mulu,htmlmulu,FileName,filefrom,htmla,htmlb,htmlc,htmld)
if mulu="" then mulu=""&SysRootDir&""
if htmlmulu="" then htmlmulu=""&SysRootDir&""
mulu=replace(mulu, "//", "/")
FilePath=Server.MapPath(mulu)&"\"&FileName
Do_Url="http://"
Do_Url=Do_Url&Request.ServerVariables("server_name")&htmlmulu&filefrom
Do_Url=Do_Url&"?"&htmla&htmlb&"&"&htmlc&htmld
strUrl=Do_Url
set objXmlHttp=Server.createObject("Microsoft.XMLHTTP")
objXmlHttp.open "GET",strUrl,false
objXmlHttp.send()
binFileData=objXmlHttp.responseBody
Set objXmlHttp=Nothing
set objAdoStream=Server.CreateObject("Adodb." & "Stream")
objAdoStream.Type=1
objAdoStream.Open()
objAdoStream.Write(binFileData)
objAdoStream.SaveToFile FilePath,2
objAdoStream.Close()
set objAdoStream=nothing
End Function
%>