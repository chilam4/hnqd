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
msg.MailServerUserName = ""&JMailUser&"" '����smtp��������֤��½�� 
msg.MailServerPassword = ""&JMailPass&"" '����smtp��������֤���� 
msg.From = ""&JMailName&"" '������ 
msg.FromName = FromName 
msg.AddRecipient ""&BizEMail&"" '�ռ��� 
msg.Subject = ""&JMailTitle&"" '���� 
msg.Body = ""&emailcontent&"" '���� 
msg.Send (""&JMailSMTP&"") 'smtp��������ַ,��˾��ַΪ218.5.72.50 
set msg = nothing
end if 
%>