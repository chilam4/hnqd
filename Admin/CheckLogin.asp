<!--#include file="../Include/Const.asp"-->
<!--#include file="../Include/ConnSiteData.asp"-->
<!--#include file="../Include/Md5.asp"-->
<%
dim LoginName,LoginPassword,AdminName,Password,AdminPurview,Working,UserName,rs,sql
LoginName=trim(request.form("UserName"))
LoginPassword=Md5(request.form("password"))
CheckCode = LCase(Trim(Request("CheckCode")))
AdminLoginCode = Trim(Request("AdminLoginCode"))
If CheckCode = "" Then
   response.write "<script language='JavaScript'>alert('��֤�벻��Ϊ�գ��뷵�ؼ�飡');" & "history.back()" & "</script>"
   response.end
End If
If Trim(Session("CheckCode")) = "" Then
   response.write "<script language='JavaScript'>alert('���ڹ����¼ͣ����ʱ�������������֤��ʧЧ��\n�����·��ص�¼ҳ����е�¼��');" & "history.back()" & "</script>"
   response.end
End If
If CheckCode <> Session("CheckCode") Then
   response.write "<script language='JavaScript'>alert('���������֤���ϵͳ�����Ĳ�һ�£����������롣');" & "history.back()" & "</script>"
   response.end
End If
If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
   response.write "<script language='JavaScript'>alert('������ĺ�̨������֤�벻�ԣ����������롣');" & "history.back()" & "</script>"
   response.end
End If
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Admin where AdminName='"&LoginName&"'"
rs.open sql,conn,1,3
if rs.eof then
   response.write "<script language='JavaScript'>alert('�޴��ʺţ��뷵�ؼ�飡');" & "history.back()" & "</script>"
   response.end
else
   AdminName=rs("AdminName")
   Password=rs("Password")
   AdminPurview=rs("AdminPurview")
   Working=rs("Working")
   UserName=rs("UserName")
end if
if LoginPassword<>Password then
   response.write "<script language='JavaScript'>alert('��������뷵�ؼ�飡');" & "history.back()" & "</script>"
   response.end
end if
if not Working then
   response.write "<script language='JavaScript'>alert('�ʺű����ã�');" & "history.back()" & "</script>"
   response.end
end if
if LoginName=AdminName and LoginPassword=Password then
   rs("LastLoginTime")=now()
   rs("LastLoginIP")=Request.ServerVariables("Remote_Addr")
   rs.update
   rs.close
   set rs=nothing
   session("AdminName")=AdminName
   session("UserName")=UserName
   session("AdminPurview")=AdminPurview
   session("LoginSystem")="Succeed"
   session.timeout=60
   Response.Cookies("AdminLoginCode")=AdminLoginCode
   dim LoginIP,LoginTime,LoginSoft
   LoginIP=Request.ServerVariables("Remote_Addr")
   LoginSoft=Request.ServerVariables("Http_USER_AGENT")
   LoginTime=now()
   set rs = server.createobject("adodb.recordset")
   sql="select * from Qianbo_AdminLog"
   rs.open sql,conn,1,3
   rs.addnew
   rs("AdminName")=AdminName
   rs("UserName")=UserName
   rs("LoginIP")=LoginIP
   rs("LoginSoft")=LoginSoft
   rs("LoginTime")=LoginTime
   rs.update
   rs.close
   set rs=nothing
   response.redirect "Admin_Index.asp"
   response.end
end if
%>