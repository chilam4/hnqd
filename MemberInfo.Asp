<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>�޸Ļ�Ա���� - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then
  response.Write "<script language=javascript>alert('���ȵ�¼��');location.href='Index.asp';</script>"
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
                <td width="200" align="right">�û����ƣ�</td>
                <td><% =mMemName %></td>
              </tr>
              <tr>
                <td width="200" align="right">�������</td>
                <td><% =mGroupIdName %></td>
              </tr>
              <tr>
                <td align="right">��ʵ������</td>
                <td><input name="RealName" type="text" id="RealName" size="20" maxlength="50" value="<%=mRealName%>" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�Ա�</td>
                <td><input name="Sex" type="radio" value="����" class="inputnoborder" <%if mSex="����" then response.write ("checked")%> />���� <input type="radio" name="Sex" value="Ůʿ" class="inputnoborder" <%if mSex="Ůʿ" then response.write ("checked")%> />Ůʿ <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�û����룺</td>
                <td><input name="Password" type="password" id="Password" size="20" maxlength="16" /> <font color="red">���޸�������</font></td>
              </tr>
              <tr>
                <td align="right">�ظ��û����룺</td>
                <td><input name="vPassword" type="password" id="vPassword" size="20" maxlength="16" /></td>
              </tr>
              <tr>
                <td align="right">���뱣�����⣺</td>
                <td><% =mQusetion %></td>
              </tr>
              <tr>
                <td align="right">���뱣���𰸣�</td>
                <td>������(�����޸�)</td>
              </tr>
              <tr>
                <td align="right">��λ���ƣ�</td>
                <td><input name="Company" type="text" id="Company" value="<% =mCompany %>" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ϸ��ַ��</td>
                <td><input name="Address" type="text" id="Address" value="<% =mAddress %>" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�������룺</td>
                <td><input name="ZipCode" type="text" id="ZipCode" value="<% =mZipCode %>" size="20" maxlength="20" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ϵ�绰��</td>
                <td><input name="Telephone" type="text" id="Telephone" value="<% =mTelephone %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">������룺</td>
                <td><input name="Fax" type="text" id="Fax" value="<% =mFax %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�ֻ����룺</td>
                <td><input name="Mobile" type="text" id="Mobile" value="<% =mMobile %>" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�������䣺</td>
                <td><input name="Email" type="text" id="Email" value="<% =mEmail %>" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ַ��</td>
                <td><input name="HomePage" type="text" id="HomePage" value="<% =mHomePage %>" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">ע��ʱ�䣺</td>
                <td><% =mAddTime %></td>
              </tr>
              <tr>
                <td align="right">��¼������</td>
                <td><% =mLoginTimes %> ��</td>
              </tr>
              <tr>
                <td align="right">����¼ʱ�䣺</td>
                <td><% =mLastLoginTime %></td>
              </tr>
              <tr>
                <td align="right">����¼IP��</td>
                <td><% =mLastLoginIP %></td>
              </tr>
              <tr>
                <td align="right">��֤�룺</td>
                <td><input name="CheckCode" type="text" size="6" maxlength="6" /> <a href="javascript:refreshimg()" title="�������������ͼƬ��"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input name="Submit" type="submit" value="�޸Ļ�Ա����" /></td>
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
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""MemberCenter.asp"" class=""agray"">��Ա����</a> - �û�����"&VbCrLf
end Function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>���������Ϣ</center>"
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