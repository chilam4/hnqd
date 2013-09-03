<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>修改会员资料 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then
  response.Write "<script language=javascript>alert('请先登录！');location.href='Index.asp';</script>"
  response.end
end if
dim MemID,mMemName,mGroupIdName,mRealName,mSex,mQusetion
dim mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,mHomepage,mAddTime,mLoginTimes,mLastLoginTime,mLastLoginIP
call MemInfo()
%>
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
          <td height="20"><table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
		  <form action="MemberSaveInfo.asp?ID=<%=MemID%>" method="post" name="formInfo" id="formInfo">
              <tr>
                <td width="200" align="right">用户名称：</td>
                <td><% =mMemName %></td>
              </tr>
              <tr>
                <td width="200" align="right">所属组别：</td>
                <td><% =mGroupIdName %></td>
              </tr>
              <tr>
                <td align="right">真实姓名：</td>
                <td><input name="RealName" type="text" id="RealName" size="20" maxlength="50" value="<%=mRealName%>" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">性别：</td>
                <td><input name="Sex" type="radio" value="先生" class="inputnoborder" <%if mSex="先生" then response.write ("checked")%> />先生 <input type="radio" name="Sex" value="女士" class="inputnoborder" <%if mSex="女士" then response.write ("checked")%> />女士 <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">用户密码：</td>
                <td><input name="Password" type="password" id="Password" size="20" maxlength="16" /> <font color="red">不修改请留空</font></td>
              </tr>
              <tr>
                <td align="right">重复用户密码：</td>
                <td><input name="vPassword" type="password" id="vPassword" size="20" maxlength="16" /></td>
              </tr>
              <tr>
                <td align="right">密码保护问题：</td>
                <td><% =mQusetion %></td>
              </tr>
              <tr>
                <td align="right">密码保护答案：</td>
                <td>已隐藏(不可修改)</td>
              </tr>
              <tr>
                <td align="right">单位名称：</td>
                <td><input name="Company" type="text" id="Company" value="<% =mCompany %>" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">详细地址：</td>
                <td><input name="Address" type="text" id="Address" value="<% =mAddress %>" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">邮政编码：</td>
                <td><input name="ZipCode" type="text" id="ZipCode" value="<% =mZipCode %>" size="20" maxlength="20" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">联系电话：</td>
                <td><input name="Telephone" type="text" id="Telephone" value="<% =mTelephone %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">传真号码：</td>
                <td><input name="Fax" type="text" id="Fax" value="<% =mFax %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">手机号码：</td>
                <td><input name="Mobile" type="text" id="Mobile" value="<% =mMobile %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">电子信箱：</td>
                <td><input name="Email" type="text" id="Email" value="<% =mEmail %>" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">网址：</td>
                <td><input name="HomePage" type="text" id="HomePage" value="<% =mHomePage %>" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">注册时间：</td>
                <td><% =mAddTime %></td>
              </tr>
              <tr>
                <td align="right">登录次数：</td>
                <td><% =mLoginTimes %> 次</td>
              </tr>
              <tr>
                <td align="right">最后登录时间：</td>
                <td><% =mLastLoginTime %></td>
              </tr>
              <tr>
                <td align="right">最后登录IP：</td>
                <td><% =mLastLoginIP %></td>
              </tr>
              <tr>
                <td align="right">验证码：</td>
                <td><input name="CheckCode" type="text" size="6" maxlength="6" /> <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input name="Submit" type="submit" value="修改会员资料" /></td>
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
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MemberCenter.asp"" class=""agray"">会员中心</a> - 用户资料"&VbCrLf
end Function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>暂无相关信息</center>"
  else
    MemID=rs("ID")
	mMemName=rs("MemName")
	mGroupIdName=GroupName(rs("GroupID"))
	mRealName=rs("RealName")
	mSex=rs("Sex")
	mQusetion=rs("Question")
	mCompany=rs("Company")
	mAddress=rs("Address")
	mZipCode=rs("ZipCode")
	mTelephone=rs("Telephone")
	mFax=rs("Fax")
	mMobile=rs("Mobile")
	mEmail=rs("Email")
	mHomePage=rs("HomePage")
	mAddTime=rs("AddTime")
	mLoginTimes=rs("LoginTimes")
	mLastLoginTime=rs("LastLoginTime")
	mLastLoginIP=rs("LastLoginIP")
  end if
  rs.close
  set rs=nothing
end sub

function GroupName(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupName=rs("GroupName")
  rs.close
  set rs=nothing
end function
%>
<script language="javascript">
<!--
function refreshimg(){document.all.checkcode.src="Include/CheckCode/CheckCode.asp";}
//-->
</script>