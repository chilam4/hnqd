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
ID=request.QueryString("ID")
if ID <> "" or IsNumeric(ID) then
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Others where ViewFlag and ID="&ID
rs.open sql,conn,1,3
If rs("SeoKeywords") <> "" Then
SeoKeywords=rs("SeoKeywords")
Else
SeoKeywords=rs("OthersName")
End If
If rs("SeoDescription") <> "" Then
SeoDescription=rs("SeoDescription")
Else
SeoDescription=rs("OthersName")
End If
SeoTitle=rs("OthersName")
End If
rs.close
set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
<meta name="keywords" content="<% =SeoKeywords %>" />
<meta name="description" content="<% =SeoDescription %>" />
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
          <td bgcolor="#FFFFFF"><div id="nav"><%=WebMenu(29,0,0)%></div></td>
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
        <tr>
          <td><table width="733" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="30" height="36"><img src="images/title_bg1.gif" width="30" height="36" /></td>
                <td width="673" align="center" background="images/title_bg_center.jpg" style="color:#000000; font-size:14px; font-weight:bold">
<%
  ID=request.QueryString("ID")
  if ID="" or not IsNumeric(ID) then
    response.write "<center>暂无相关信息</center>"
  elseif conn.execute("select * from Qianbo_Others Where ViewFlag and  ID="&ID).eof then
    response.write "<center>暂无相关信息</center>"
  else
	dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_Others where ViewFlag and ID="&ID
    rs.open sql,conn,1,3
	Response.Write ""&rs("OthersName")&""&VbCrLf
  End If
%></td>
                <td width="30"><img src="images/title_bg2.gif" width="30" height="36" /></td>
              </tr>
              <tr>
                <td><img src="images/title_bg3.gif" width="30" height="14" /></td>
                <td background="images/title_bg_center2.jpg">&nbsp;</td>
                <td><img src="images/title_bg4.gif" width="30" height="14" /></td>
              </tr>
            </table></td>
        </tr>
<%
if ViewNoRight(rs("GroupID"),rs("Exclusive")) Then
Response.Write "        <tr>"&VbCrLf
Response.Write "          <td align=""center"" style=""color:#999999""><strong>更新时间：</strong>"&FormatDate(rs("UpdateTime"),13)&" <strong>浏览次数：</strong><script language=""javascript"" src=""HitCount.asp?id="&rs("ID")&"&LX=Qianbo_Others""></script><script language=""javascript"" src=""HitCount.asp?action=count&LX=Qianbo_Others&id="&rs("ID")&"""></script></td>"&VbCrLf
Response.Write "        </tr>"&VbCrLf
Response.Write "        <tr>"&VbCrLf
Response.Write "          <td height=""20"">&nbsp;</td>"&VbCrLf
Response.Write "        </tr>"&VbCrLf
Response.Write "        <tr>"&VbCrLf
Response.Write "          <td style=""line-height:25px; font-size:12px; padding-left: 8px; padding-right: 8px"">"&rs("Content")&"</td>"&VbCrLf
Response.Write "        </tr>"&VbCrLf
Else
Response.Write "        <tr>"&VbCrLf
Response.Write "          <td align=""center""><img src=""Images/NoRight.jpg""></td>"&VbCrLf
Response.Write "        </tr>"&VbCrLf
end if
rs.update
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
function WebMenu(ParentID,i,level)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_OthersSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
  rs.open sql,conn,1,1
  if conn.execute("select ID from Qianbo_OthersSort Where ViewFlag and ParentID=0").eof then
    response.write "<center>暂无相关信息</center>"
  end if
  do while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&OtherSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "hyList.asp?SortID="&rs("ID")&""
	End If
	response.write "<a href="""&AutoLink&""" class=""navA"">"&rs("SortName")&"</a>"&VbCrLf
    i=i+1
	if i<level then call WebMenu(rs("ID"),i,level)
	i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end Function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""hyList.asp"" class=""agray"">行业门户服务</a>"&VbCrLf
  if request.QueryString("ID")="" then
    WebLocation=WebLocation
  elseif not IsNumeric(request.QueryString("ID")) then
    WebLocation=WebLocation&"读取错误"
  elseif conn.execute("select * from Qianbo_Others Where ViewFlag and  ID="&request.QueryString("ID")).eof then
    WebLocation=WebLocation&"读取错误"
  else
    dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_Others where ViewFlag and ID="&request.QueryString("ID")
    rs.open sql,conn,1,1
	WebLocation=WebLocation&SortPathTXT("Qianbo_OthersSort",rs("SortID"))
    rs.close
    set rs=nothing
  end if
end Function

function SortPathTXT(DataFrom,ID)
  dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "&DataFrom&" where ViewFlag and ID="&ID
  rs.open sql,conn,1,1
  if not rs.eof Then
	If ISHTML = 1 Then
		AutoLink = ""&OtherSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "hyList.asp?SortID="&rs("ID")&""
	End If
	SortPathTXT=SortPathTXT(DataFrom,rs("ParentID"))&" - <a href="""&AutoLink&"""><span style=""color:#EC0000"">"&rs("SortName")&"</span></a>"
  end if
  rs.close
  set rs=nothing
end Function
%>