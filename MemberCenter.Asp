<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>��Ա���� - <% =SiteTitle %></title>
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
    <td width="773" valign="top"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
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
        <% if session("MemName")="" or session("GroupID")="" or session("MemLogin")<>"Succeed" then %>
        <tr>
          <td><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr style="line-height:25px; font-size:12px" align="center">
                <td><img src="images/Page_error.jpg" align="absmiddle" /></td>
                <td>��л���Ĺ�ע���������ݣ�<font color="red">�ο�</font><br />
                  <a href="MemberRegister.asp"><font color="red">����ע���Ա</font></a>���Բ鿴�����Ʒ��Ѷ��</td>
              </tr>
            </table></td>
        </tr>
        <% else %>
        <tr>
          <td><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr style="line-height:25px; font-size:12px" align="center">
                <td><img src="images/Page_ok.jpg" align="absmiddle" /></td>
                <td>��л���Ĺ�ע���������ݣ�<font color="red"><%=MemGroup(session("GroupID"))%></font><br />
                  ��ѡ�������˵������������Ļ�Ա���ϡ����������ԡ�ӦƸ����Ϣ��</td>
              </tr>
            </table></td>
        </tr>
        <% end if %>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
Function WebMenu()
If session("MemName")="" Or session("GroupID")="" Or session("MemLogin")<>"Succeed" Then
  response.write "<a href=""MemberRegister.asp"">ע���Ա</a>"
Else
  response.write "<a href=""MemberInfo.asp"">�޸�ע������</a>"
  response.write "<a href=""MemberMessage.asp"">�ҵ�����</a>"
  response.write "<a href=""MemberOrder.asp"">�ҵĶ���</a>"
  response.write "<a href=""MemberTalent.asp"">�ҵ�ӦƸ</a>"
  response.write "<a href=""MemberLogin.asp?Action=Out"">�˳���¼</a>"
End If
End Function

function WebLocation()
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""MemberCenter.asp"" class=""agray"">��Ա����</a> "&VbCrLf
end function

function MemGroup(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  MemGroup=rs("GroupName")
  rs.close
  set rs=nothing
end function
%>