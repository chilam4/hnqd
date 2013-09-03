<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
BizContent=Request("contenthide")
BizPhone=Request("phonehide")
BizEMail=Request("emailhide")
BizAddr=Request("addrhide")
emailcontent=BizContent

set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Biz"
rs.open sql,conn,1,3
rs.addnew
rs("BizContent")=BizContent
rs("BizPhone")=BizPhone
rs("BizEMail")=BizEMail
rs("BizAddr")=BizAddr
rs("BizDate")=now()
rs("BizOK")=0
rs.update
rs.close
set rs=nothing
%>
<%
if JMailPubDisplay="1" then
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName = ""&JMailUser&"" '输入smtp服务器验证登陆名 
msg.MailServerPassword = ""&JMailPass&"" '输入smtp服务器验证密码 
msg.From = ""&JMailName&"" '发件人 
msg.FromName = FromName 
msg.AddRecipient ""&BizEMail&"" '收件人 
msg.Subject = ""&JMailTitle&"" '主题 
msg.Body = ""&emailcontent&"" '正文 
msg.Send (""&JMailSMTP&"") 'smtp服务器地址,本司地址为218.5.72.50 
set msg = nothing
end if 
%>