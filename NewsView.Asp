<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
Call SiteInfo()
If ISHTML = 1 Then
    Response.expires = 0
    Response.expiresabsolute = Now() - 1
    Response.addHeader "pragma", "no-cache"
    Response.addHeader "cache-control", "private"
    Response.CacheControl = "no-cache"
End If
ID = request.QueryString("ID")
If ID <> "" Or IsNumeric(ID) Then
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_News where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 3
    If rs("SeoKeywords") <> "" Then
        SeoKeywords = rs("SeoKeywords")
    Else
        SeoKeywords = rs("NewsName")
    End If
    If rs("SeoDescription") <> "" Then
        SeoDescription = rs("SeoDescription")
    Else
        SeoDescription = rs("NewsName")
    End If
    SeoTitle = rs("NewsName")
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
<meta name="keywords" content="<% =SeoKeywords %>" />
<meta name="description" content="<% =SeoDescription %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div id="wrapper">
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background:url(images/main_bg.jpg) repeat-x;">
  <tr>
    <td width="230" align="left" valign="top" style="background:url(images/left_bg.gif) repeat-y right;"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
      <tr>
        <td width="190"><img src="images/left_nav.gif" width="190" height="28" /></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><div id="nav"><%=WebMenu(12,0,0)%></div></td>
      </tr>
      <tr>
        <td height="18">&nbsp;</td>
      </tr>
    </table>
      <!--#include file="Center_Left.asp" --></td>
    <td width="773" valign="top"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
      <tr>
        <td width="733" style="background:url(images/bg1.gif) repeat-x left bottom;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%" align="right" style="color:#999999"><img src="images/arr6.gif" width="9" height="9" align="absmiddle" />面包屑</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="18">&nbsp;</td>
      </tr>
      <tr>
        <td align="center">暂无相关信息</td>
      </tr>
      <tr>
        <td><table width="733" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="30" height="36"><img src="images/title_bg1.gif" width="30" height="36" /></td>
            <td width="673" align="center" background="images/title_bg_center.jpg" style="color:#000000; font-size:14px; font-weight:bold">文章标题</td>
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
        <td align="center" style="color:#999999">
        更新时间
        </td>
      </tr>
      <tr>
        <td height="20">&nbsp;</td>
      </tr>
      <tr>
        <td style="line-height:25px; font-size:12px; padding-left: 8px; padding-right: 8px">我是内容</td>
      </tr>
      <tr>
        <td align="center"><img src="Images/NoRight.jpg" /></td>
      </tr>
      <tr>
        <td height="20">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<!--#include file="End.asp" -->
</div>
</body>
</html>
<%
Function WebMenu(ParentID, i, level)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_NewsSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
    rs.Open sql, conn, 1, 1
    If conn.Execute("select ID from Qianbo_NewsSort Where ViewFlag and ParentID=0").EOF Then
        response.Write "<center>暂无相关信息</center>"
    End If
    Do While Not rs.EOF
        If ISHTML = 1 Then
            AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
        End If
        response.Write "<li><a href="""&AutoLink&""">"&rs("SortName")&"</a></li>"&vbCrLf
        i = i + 1
        If i<level Then Call WebMenu(rs("ID"), i, level)
        i = i -1
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
End Function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""About.asp"" class=""agray"">关于我们</a> - "&VbCrLf
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  if request.QueryString("ID")="" or not IsNumeric(request.QueryString("ID")) then
	sql="select top 1 * from Qianbo_About where ViewFlag and not ChildFlag order by Sequence asc"
  else
	sql="select * from Qianbo_About where ViewFlag and ID="&request.QueryString("ID")
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