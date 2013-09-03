<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>注册会员 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background-color:#fff;">
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
          <td height="20"><table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
		  <form action="MemberSaveReg.asp" method="post" name="formReg" id="formReg">
              <tr>
                <td width="200" align="right">用户名称：</td>
                <td><input name="MemName" type="text" id="MemName" size="20" maxlength="16" /> <font color="red">*</font>(0-9,a-z,-_任意组合的3-16个字符)</td>
              </tr>
              <tr>
                <td align="right">真实姓名：</td>
                <td><input name="RealName" type="text" id="RealName" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">性别：</td>
                <td><input name="Sex" type="radio" value="先生" checked="checked" class="inputnoborder" />先生 <input type="radio" name="Sex" value="女士" class="inputnoborder" />女士 <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">用户密码：</td>
                <td><input name="Password" type="password" id="Password" size="20" maxlength="16" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">重复用户密码：</td>
                <td><input name="vPassword" type="password" id="vPassword" size="20" maxlength="16" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">密码保护问题：</td>
                <td><input name="Question" type="text" id="Question" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">密码保护答案：</td>
                <td><input name="Answer" type="text" id="Answer" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">单位名称：</td>
                <td><input name="Company" type="text" id="Company" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">详细地址：</td>
                <td><input name="Address" type="text" id="Address" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">邮政编码：</td>
                <td><input name="ZipCode" type="text" id="ZipCode" size="20" maxlength="20" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">联系电话：</td>
                <td><input name="Telephone" type="text" id="Telephone" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">传真号码：</td>
                <td><input name="Fax" type="text" id="Fax" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">手机号码：</td>
                <td><input name="Mobile" type="text" id="Mobile" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">电子信箱：</td>
                <td><input name="Email" type="text" id="Email" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">网址：</td>
                <td><input name="HomePage" type="text" id="HomePage" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">验证码：</td>
                <td><input name="CheckCode" type="text" size="6" maxlength="6" /> <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input name="Submit" type="submit" value="立即注册会员，获取最新产品资讯" /></td>
              </tr>
			  </form>
            </table></td>
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
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MemberCenter.asp"" class=""agray"">会员中心</a> - 帐户注册"&VbCrLf
end function
%>
<script language="javascript">
<!--
function refreshimg(){document.all.checkcode.src="Include/CheckCode/CheckCode.asp";}
//-->
</script>