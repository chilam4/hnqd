<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>密码找回 - <% =SiteTitle %></title>
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
  response.write "<a href=""MemberRegister.asp"">注册会员</a>"
Else
  response.write "<a href=""MemberInfo.asp"">修改注册资料</a>"
  response.write "<a href=""MemberMessage.asp"">我的留言</a>"
  response.write "<a href=""MemberOrder.asp"">我的订单</a>"
  response.write "<a href=""MemberTalent.asp"">我的应聘</a>"
  response.write "<a href=""MemberLogin.asp?Action=Out"">退出登录</a>"
End If
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MemberCenter.asp"" class=""agray"">会员中心</a> - 密码找回"&VbCrLf
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
    <td width="200" align="right">会员名：</td>
    <td><input name="MemName" type="text" id="MemName" size="30" maxlength="16" /></td>
  </tr>
  <tr>
    <td align="right">电子信箱：</td>
    <td><input name="Email" type="text" id="Email" size="30" maxlength="50" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input name="Next2" type="submit" value="下一步" /></td>
  </tr>
  </form>
</table>
<%
  elseif request.QueryString("Step")="EnterNameMail" then
    set rs = server.createobject("adodb.recordset")
    sql="select * from Qianbo_Members where MemName='"&request.form("MemName")&"' and Email='"&request.form("Email")&"'"
	rs.open sql,conn,1,1
	if rs.eof then
	  WriteMsg("会员名或电子信箱地址错误！")
      exit function
	end if
%>
<table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
<form action="MemberGetPass.asp?Step=EnterAnswer&ID=<%=rs("ID")%>" method="post" name="form" id="form">
  <tr>
    <td width="200" align="right">会员名：</td>
    <td><%=rs("MemName")%></td>
  </tr>
  <tr>
    <td align="right">电子信箱：</td>
    <td><%=rs("Email")%></td>
  </tr>
  <tr>
    <td align="right">密码保护问题：</td>
    <td><%=rs("Question")%></td>
  </tr>
  <tr>
    <td align="right">密码保护答案：</td>
    <td><input name="Answer" type="text" size="30" maxlength="100" /></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input name="Next" type="submit" value="下一步" /></td>
  </tr>
  </form>
</table>
<%
  elseif request.QueryString("Step")="EnterAnswer" Then
    If trim(request("Answer"))="" Then
	WriteMsg("请填写密码保护答案！")
	exit function
	End If
    set rs = server.createobject("adodb.recordset")
    sql="select * from Qianbo_Members where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	if rs.eof then
	  WriteMsg("参数错误！")
      exit function
	end if
    if rs("Answer")<>MD5(trim(request("Answer"))) then
      WriteMsg("密码保护答案错误！")
      exit function
    end if
    randomize timer
    NewPassword=Int(899999*Rnd() +100000)
    rs("Password")=MD5(NewPassword)
    rs.update
    rs.close
    set rs=nothing
	Response.Write "<script language=javascript>alert('恭喜！您的密码找回成功！\n新密码为："&NewPassword&"\n请登录后修改为新的密码！');window.location.href='Index.asp';</script>"
    exit function
  end if
end function
%>