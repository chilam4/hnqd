<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
call SiteInfo
if ISHTML = 1 then
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
End If
if request.QueryString("ID")="" then
SeoTitle="展商服务"
elseif not IsNumeric(request.QueryString("ID")) then
SeoTitle="参数错误"
elseif conn.execute("select * from Qianbo_Zsfw Where ViewFlag and  ID="&request.QueryString("ID")).eof then
SeoTitle="参数错误"
else
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Zsfw where ViewFlag and ID="&request.QueryString("ID")
rs.open sql,conn,1,1
SeoTitle=rs("AboutName")
rs.close
set rs=nothing
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background:url(images/main_bg.jpg) repeat-x;">
  <tr>
    <td width="230" align="left" valign="top" style="background:url(images/left_bg.gif) repeat-y right;"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="190"><img src="images/left_nav.gif" width="190" height="28" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><div id="nav"><%=WebMenu()%></div></td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
      </table>
      <!--#include file="Center_Left.asp" --></td>
    <td width="773" valign="top" style="background-color:#fff;"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="733" style="background:url(images/bg1.gif) repeat-x left bottom;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" align="right" style="color:#999999"><img src="images/arr6.gif" width="9" height="9" align="absmiddle" /><%=WebLocation()%></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
<%
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  if request.QueryString("ID")="" or not IsNumeric(request.QueryString("ID")) then
	sql="select top 1 * from Qianbo_Zsfw where ViewFlag and not ChildFlag order by Sequence asc"
  else
	sql="select * from Qianbo_Zsfw where ViewFlag and ID="&request.QueryString("ID")
  end if
  rs.open sql,conn,1,3
  if rs.eof Then
%>
        <tr>
          <td align="center">暂无相关信息</td>
        </tr>
<%
  Else
	if ViewNoRight(rs("GroupID"),rs("Exclusive")) then
%>
        <tr>
          <td><table width="733" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="30" height="36"><img src="images/title_bg1.gif" width="30" height="36" /></td>
                <td width="673" align="center" background="images/title_bg_center.jpg" style="color:#000000; font-size:14px; font-weight:bold"><%=rs("AboutName")%></td>
                <td width="30"><img src="images/title_bg2.gif" width="30" height="36" /></td>
              </tr>
              <tr>
                <td><img src="images/title_bg3.gif" width="30" height="14" /></td>
                <td background="images/title_bg_center2.jpg">&nbsp;</td>
                <td><img src="images/title_bg4.gif" width="30" height="14" /></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td align="center" style="color:#999999"><strong>更新时间：</strong><%=FormatDate(rs("UpdateTime"),13)%> <strong>浏览次数：</strong><script language="javascript" src="HitCount.asp?id=<%=rs("ID")%>&LX=Qianbo_Zsfw"></script><script language="javascript" src="HitCount.asp?action=count&LX=Qianbo_Zsfw&id=<%=rs("ID")%>"></script></td>
        </tr>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
        <tr>
          <td style="line-height:25px; font-size:12px; padding-left: 8px; padding-right: 8px"><%=rs("Content")%></td>
        </tr>
<%
	else
%>
        <tr>
          <td align="center"><img src="Images/NoRight.jpg"></td>
        </tr>
<%
	end if
	rs.update
  end if
  rs.close
  set rs=nothing
%>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select ID,AboutName from Qianbo_Zsfw where ViewFlag and not ChildFlag order by Sequence asc"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>暂无相关信息</center>"
  else
    Do
		If ISHTML = 1 Then
			AutoLink = ""&ZsfwNameDiy&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "Zsfw.asp?ID="&rs("ID")&""
		End If
		response.write "<a href="""&AutoLink&""">"&rs("AboutName")&"</a>"&VbCrLf
		rs.movenext
    loop until rs.eof
  end if
  rs.close
  set rs=nothing
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""Zsfw.asp"" class=""agray"">展会概况</a> - "&VbCrLf
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  if request.QueryString("ID")="" or not IsNumeric(request.QueryString("ID")) then
	sql="select top 1 * from Qianbo_Zsfw where ViewFlag and not ChildFlag order by Sequence asc"
  else
	sql="select * from Qianbo_Zsfw where ViewFlag and ID="&request.QueryString("ID")
  end if
  rs.open sql,conn,1,1
  if rs.eof then
    WebLocation=WebLocation&"参数错误"
  else
    WebLocation=WebLocation&""&rs("AboutName")
  end if
  rs.close
  set rs=nothing
end Function
%>