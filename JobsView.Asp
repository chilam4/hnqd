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
elseif conn.execute("select * from Qianbo_Jobs Where ViewFlag and  ID="&request.QueryString("ID")).eof then
SeoTitle="参数错误"
else
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Jobs where ViewFlag and ID="&request.QueryString("ID")
rs.open sql,conn,1,1
SeoTitle=rs("JobName")
rs.close
set rs=nothing
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>招聘信息：<% =SeoTitle %> - <% =SiteTitle %></title>
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
    <td width="773" valign="top" style="background-color:#ffffff;"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
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
          <td height="20"><%=WebContent("Qianbo_Jobs",request.QueryString("ID"))%></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu()
  response.write "<a href=""JobsList.asp"">招聘信息</a>"
  response.write "<a href=""MemberTalent.asp"">我的应聘</a>"
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""JobsList.asp"" class=""agray"">招聘信息</a>"&VbCrLf
end Function

function WebContent(DataFrom,ID)
  if ID="" or not IsNumeric(ID) then
    response.write "<div class=""Bodyer_right_page_content"">信息ID读取错误</div>"
  elseif conn.execute("select * from "&DataFrom&" Where ViewFlag and  ID="&ID).eof then
    response.write "<div class=""Bodyer_right_page_content"">信息ID读取错误</div>"
  else
	dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from "&DataFrom&" where ViewFlag and ID="&ID
    rs.open sql,conn,1,3
	Response.Write "<table width=""95%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""5"">"&VbCrLf
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td style=""background:url(images/bg1.gif) repeat-x left bottom""><strong style=""color:red"">"&rs("JobName")&"</strong>(更新时间："&FormatDate(rs("UpdateTime"),13)&" 浏览次数：<script language=""javascript"" src=""HitCount.asp?id="&rs("ID")&"&LX=Qianbo_Jobs""></script><script language=""javascript"" src=""HitCount.asp?action=count&LX=Qianbo_Jobs&id="&rs("ID")&"""></script>)</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td style=""line-height:200%"">工作地点："&rs("JobAddress")&"<br />"&VbCrLf
	Response.Write "      招聘人数："&rs("JobNumber")&"人<br />"&VbCrLf
	Response.Write "      薪酬待遇："&rs("Emolument")&"<br />"&VbCrLf
	Response.Write "      有效时间："&FormatDate(rs("StartDate"),13)&" 至 "&FormatDate(rs("EndDate"),13)&"<br />"&VbCrLf
	Response.Write "      招聘单位："&rs("eEmployer")&"<br />"&VbCrLf
	Response.Write "      联 系 人："&rs("eContact")&"<br />"&VbCrLf
	Response.Write "      联系电话："&rs("eTel")&"<br />"&VbCrLf
	Response.Write "      详细地址："&rs("eAddress")&"<br />"&VbCrLf
	Response.Write "	  邮政编码："&rs("ePostCode")&"<br />"&VbCrLf
	Response.Write "      电子邮箱："&rs("eEmail")&"<br />"&VbCrLf
	Response.Write "      <br />"&VbCrLf
	Response.Write "      <font color=""#ff3300"">工作职责：</font><br />"&VbCrLf
	Response.Write "      "&HtmlStrReplace(rs("Responsibility"))&"<br />"&VbCrLf
	Response.Write "      <br />"&VbCrLf
	Response.Write "      <font color=""#ff3300"">职位要求：</font><br />"&VbCrLf
	Response.Write "      "&HtmlStrReplace(rs("Requirement"))&"<br />"&VbCrLf
	Response.Write "      <br />"&VbCrLf
	Response.Write "      <a href=""TalentWrite.asp?JobID="&ID&"&JobName="&server.urlencode(rs("JobName"))&""" title=""提交简历"" ><img src=""Images/Page_resume.gif"" border=""0"" /></a></td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	Response.Write "</table><br />"&VbCrLf
	rs.update
    rs.close
    set rs=nothing
  end if
end function
%>