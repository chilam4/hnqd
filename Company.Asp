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
SeoTitle="关于我们"
elseif not IsNumeric(request.QueryString("ID")) then
SeoTitle="参数错误"
elseif conn.execute("select * from Qianbo_About Where ViewFlag and  ID="&request.QueryString("ID")).eof then
SeoTitle="参数错误"
else
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_About where ViewFlag and ID="&request.QueryString("ID")
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
</head>
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background-color:#fff;">
  <tr>
    <td width="230" align="left" valign="top"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
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
    <td width="773" valign="top"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="733" style="background:url(images/bg1.gif) repeat-x left bottom;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" align="right" style="color:#999999"><img src="images/arr6.gif" width="9" height="9" align="absmiddle" /><%=WebLocation()%></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
        <tr>
          <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="213" valign="top" style="background:url(images/bg-line.gif) repeat-y right top;"><img src="images/us.jpg" width="200" height="260" /></td>
                <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-left:18px;">
                    <tr>
                      <td><table width="90%" border="1" bordercolor="#EDEDED" cellspacing="0" cellpadding="8" style="border-collapse: collapse">
                          <tr>
                            <td bgcolor="#EDEDED" style="font-weight: bold; color: #EC0000"><%=SiteTitle%></td>
                          </tr>
                          <tr>
                            <td>公司名称：<%=ComName%></td>
                          </tr>
                          <tr>
                            <td>公司地址：<%=Address%></td>
                          </tr>
                          <tr>
                            <td>邮政编码：<%=ZipCode%></td>
                          </tr>
                          <tr>
                            <td>联系电话：<%=Telephone%> 传真号码：<%=Fax%></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
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
  sql="select ID,AboutName from Qianbo_About where ViewFlag and not ChildFlag order by Sequence asc"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>暂无相关信息</center>"
  else
    Do
		If ISHTML = 1 Then
			AutoLink = ""&AboutNameDiy&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "About.asp?ID="&rs("ID")&""
		End If
		response.write "<a href="""&AutoLink&""">"&rs("AboutName")&"</a>"&VbCrLf
		rs.movenext
    loop until rs.eof
  end if
  rs.close
  set rs=nothing
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""Company.asp"" class=""agray"">联系我们</a>"&VbCrLf
end Function
%>