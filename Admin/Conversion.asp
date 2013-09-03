<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
ID=request.QueryString("ID")
LX=request.QueryString("LX")
Operation=request.QueryString("Operation")
strReferer=Request.ServerVariables("http_referer")
If Operation = "up" Then
Conn.execute "update "&LX&" set ViewFlag = 1 where ID=" & ID
ElseIf Operation = "BizOK" Then
Conn.execute "update "&LX&" set BizOK = 1 where ID=" & ID
Else
Conn.execute "update "&LX&" set ViewFlag = 0 where ID=" & ID
End If
response.Redirect strReferer
%>