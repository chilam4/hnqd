<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>申请职位 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
dim JobID,TalentsName,BirthDate,Stature,Marriage,RegResidence,EduResume,JobResume
JobID=request.QueryString("JobID")
TalentsName=request.QueryString("JobName")
dim mMemID,mLinkman,mSex,mAddress,mZipCode,mTelephone,mMobile,mEmail
if session("MemName")<>"" and session("MemLogin")="Succeed" then
  call MemInfo()
else
  mSex="先生"
  mMemID=0
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
              <form action="TalentSave.asp?MemberID=<%=mMemID%>&JobID=<%=JobID%>" method="post" name="formWrite" id="formWrite">
                <tr>
                  <td width="200" align="right">申请职位：</td>
                  <td><input name="TalentsName" type="text" id="TalentsName" value="<%=TalentsName%>" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">姓名：</td>
                  <td><input name="Linkman" type="text" value="<%=mLinkman%>" size="20" maxlength="50" />
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
                  <td align="right">出生日期：</td>
                  <td><input name="BirthDate" type="text" id="BirthDate" size="20" maxlength="10" />
                    <font color="red">*</font>(格式如：1988-08-18)</td>
                </tr>
                <tr>
                  <td align="right">婚姻状况：</td>
                  <td><input name="Marriage" type="radio" value="未婚" checked="checked" class="inputnoborder" />
                    未婚
                    <input type="radio" name="Marriage" value="已婚" class="inputnoborder" />
                    已婚 <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">身高：</td>
                  <td><input name="Stature" type="text" id="Stature" size="20" maxlength="10" />CM
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">户口所在地：</td>
                  <td><input name="RegResidence" type="text" id="RegResidence" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">教育经历：</td>
                  <td>学历文凭&nbsp;&nbsp;&nbsp;&nbsp;起止日期&nbsp;&nbsp;&nbsp;&nbsp;专业名称&nbsp;&nbsp;&nbsp;&nbsp;相关证书&nbsp;&nbsp;&nbsp;&nbsp;就读学校 <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right"></td>
                  <td><textarea name="EduResume" rows="8" style="width:500px"></textarea></td>
                </tr>
                <tr>
                  <td align="right">工作经历：</td>
                  <td>起止日期&nbsp;&nbsp;&nbsp;&nbsp;职位名称&nbsp;&nbsp;&nbsp;&nbsp;工作内容&nbsp;&nbsp;&nbsp;&nbsp;就职单位 <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right"></td>
                  <td><textarea name="JobResume" rows="8" id="JobResume" style="width:500px"></textarea></td>
                </tr>
                <tr>
                  <td align="right">联系地址：</td>
                  <td><input name="Address" type="text" value="<%=mAddress%>" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">邮政编码：</td>
                  <td><input name="ZipCode" type="text" id="ZipCode" value="<%=mZipCode%>" size="20" maxlength="20" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">联系电话：</td>
                  <td><input name="Telephone" type="text" id="Telephone" value="<%=mTelephone%>" size="20" maxlength="50" /></td>
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
                  <td align="right">验证码：</td>
                  <td><input name="CheckCode" type="text" size="6" maxlength="6" />
                    <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td><input name="Submit" type="submit" value="提交简历" /></td>
                </tr>
              </form>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</td>
</tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu()
  response.write "<a href=""JobsList.asp"">招聘信息</a>"
  response.write "<a href=""MemberTalent.asp"">我的应聘</a>"
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - 简历提交"&VbCrLf
end Function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "<center>暂无相关信息</center>"
  else
    mMemID=rs("ID")
	if not rs("RealName")="" then
	  mLinkman=rs("RealName")
	else
	  mLinkman=rs("MemName")
	end if
	mSex=rs("Sex")
	mAddress=rs("Address")
	mZipCode=rs("ZipCode")
	mTelephone=rs("Telephone")
	mMobile=rs("Mobile")
	mEmail=rs("Email")
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