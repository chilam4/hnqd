<%
Dim Conn,ConnStr
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(""&SysRootDir&""&SiteDataPath&"")
Conn.open ConnStr
If err Then
   err.clear
   Set Conn = Nothing
   Response.Write "数据库连接错误，请检查连接参数！"
   Response.End
End If
%>
<!--#include file="Function.asp" -->