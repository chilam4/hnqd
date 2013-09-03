<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>留言反馈 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
dim MesName,Content,SecretFlag,mMemID,mLinkman,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail
if session("MemName")<>"" and session("MemLogin")="Succeed" then
  call MemInfo()
else
  mSex="先生"
  mMemID=0
  SecretFlag="<font color=""red"">会员功能，请先注册会员。</font>"
end if
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
              <form action="MessageSave.asp?MemberID=<%=mMemID%>" method="post" name="formWrite" id="formWrite">
                <tr>
                  <td width="200" align="right">留言主题：</td>
                  <td><input name="MesName" type="text" id="MesName" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">留言内容：</td>
                  <td><textarea name="Content" cols="70" rows="8"></textarea>
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">称呼：</td>
                  <td><input name="Linkman" type="text" id="Linkman" value="<%=mLinkman%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">性别：</td>
                  <td><input name="Sex" type="radio" value="先生" class="inputnoborder" <%if mSex="先生" then response.write ("checked")%> />
                    先生
                    <input type="radio" name="Sex" value="女士" class="inputnoborder" <%if mSex="女士" then response.write ("checked")%> />
                    女士 <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">公司名称：</td>
                  <td><input name="Company" type="text" value="<%=mCompany%>" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">联系地址：</td>
                  <td><input name="Address" type="text" value="<%=mAddress%>" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">邮政编码：</td>
                  <td><input name="ZipCode" type="text" value="<%=mZipCode%>" size="20" maxlength="20" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">联系电话：</td>
                  <td><input name="Telephone" type="text" id="Telephone" value="<%=mTelephone%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">传真号码：</td>
                  <td><input name="Fax" type="text" id="Fax" value="<%=mFax%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">手机号码：</td>
                  <td><input name="Mobile" type="text" id="Mobile" value="<%=mMobile%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">电子信箱：</td>
                  <td><input name="Email" type="text" id="Email" value="<%=mEmail%>" size="30" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">悄悄话：</td>
                  <td><%=SecretFlag%></td>
                </tr>
                <tr>
                  <td align="right">验证码：</td>
                  <td><input name="CheckCode" type="text" size="6" maxlength="6" />
                    <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td><input name="Submit" type="submit" value="发表留言" /></td>
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
  response.write "<a href=""MessageWrite.asp"">签写留言</a>"
  response.write "<a href=""MessageList.asp"">所有留言</a>"
  response.write "<a href=""MemberMessage.asp"">我的留言</a>"
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MessageList.asp"" class=""agray"">留言反馈</a> - 签写留言"&VbCrLf
end Function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>参数错误</center>"
  else
    mMemID=rs("ID")
	if not rs("RealName")="" then
	  mLinkman=rs("RealName")
	else
	  mLinkman=rs("MemName")
	end if
	mSex=rs("Sex")
	mCompany=rs("Company")
	mAddress=rs("Address")
	mZipCode=rs("ZipCode")
	mTelephone=rs("Telephone")
	mFax=rs("Fax")
	mMobile=rs("Mobile")
	mEmail=rs("Email")
	SecretFlag="<input type=""checkbox"" name=""SecretFlag"" value=""1"" class=""inputnoborder"">"
  end if
  rs.close
  set rs=nothing
end sub
%>
<script language="javascript">
<!--
function refreshimg(){document.all.checkcode.src="Include/CheckCode/CheckCode.asp";}
//-->
</script>