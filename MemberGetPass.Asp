<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>�����һ� - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<!--#include file="Include/Md5.asp" -->
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
        <tr>
          <td height="20"><%=WebContent()%></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu()
If session("MemName")="" Or session("GroupID")="" Or session("MemLogin")<>"Succeed" Then
  response.write "<a href=""MemberRegister.asp"">ע���Ա</a>"
Else
  response.write "<a href=""MemberInfo.asp"">�޸�ע������</a>"
  response.write "<a href=""MemberMessage.asp"">�ҵ�����</a>"
  response.write "<a href=""MemberOrder.asp"">�ҵĶ���</a>"
  response.write "<a href=""MemberTalent.asp"">�ҵ�ӦƸ</a>"
  response.write "<a href=""MemberLogin.asp?Action=Out"">�˳���¼</a>"
End If
end function

function WebLocation()
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""MemberCenter.asp"" class=""agray"">��Ա����</a> - �����һ�"&VbCrLf
end Function

function MemGroup(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  MemGroup=rs("GroupName")
  rs.close
  set rs=nothing
end function

function WebContent()
  dim rs,sql,NewPassword
  if request.QueryString("Step")="" then
%>
<table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
<form action="MemberGetPass.asp?Step=EnterNameMail" method="post" name="form" id="form">
  <tr>
    <td width="200" align="right">��Ա����</td>
    <td><input name="MemName" type="text" id="MemName" size="30" maxlength="16" /></td>
  </tr>
  <tr>
    <td align="right">�������䣺</td>
    <td><input name="Email" type="text" id="Email" size="30" maxlength="50" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input name="Next2" type="submit" value="��һ��" /></td>
  </tr>
  </form>
</table>
<%
  elseif request.QueryString("Step")="EnterNameMail" then
    set rs = server.createobject("adodb.recordset")
    sql="select * from Qianbo_Members where MemName='"&request.form("MemName")&"' and Email='"&request.form("Email")&"'"
	rs.open sql,conn,1,1
	if rs.eof then
	  WriteMsg("��Ա������������ַ����")
      exit function
	end if
%>
<table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
<form action="MemberGetPass.asp?Step=EnterAnswer&ID=<%=rs("ID")%>" method="post" name="form" id="form">
  <tr>
    <td width="200" align="right">��Ա����</td>
    <td><%=rs("MemName")%></td>
  </tr>
  <tr>
    <td align="right">�������䣺</td>
    <td><%=rs("Email")%></td>
  </tr>
  <tr>
    <td align="right">���뱣�����⣺</td>
    <td><%=rs("Question")%></td>
  </tr>
  <tr>
    <td align="right">���뱣���𰸣�</td>
    <td><input name="Answer" type="text" size="30" maxlength="100" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input name="Next" type="submit" value="��һ��" /></td>
  </tr>
  </form>
</table>
<%
  elseif request.QueryString("Step")="EnterAnswer" Then
    If trim(request("Answer"))="" Then
	WriteMsg("����д���뱣���𰸣�")
	exit function
	End If
    set rs = server.createobject("adodb.recordset")
    sql="select * from Qianbo_Members where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	if rs.eof then
	  WriteMsg("��������")
      exit function
	end if
    if rs("Answer")<>MD5(trim(request("Answer"))) then
      WriteMsg("���뱣���𰸴���")
      exit function
    end if
    randomize timer
    NewPassword=Int(899999*Rnd() +100000)
    rs("Password")=MD5(NewPassword)
    rs.update
    rs.close
    set rs=nothing
	Response.Write "<script language=javascript>alert('��ϲ�����������һسɹ���\n������Ϊ��"&NewPassword&"\n���¼���޸�Ϊ�µ����룡');window.location.href='Index.asp';</script>"
    exit function
  end if
end function
%>