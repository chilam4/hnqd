<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>���Է��� - <% =SiteTitle %></title>
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
  mSex="����"
  mMemID=0
  SecretFlag="<font color=""red"">��Ա���ܣ�����ע���Ա��</font>"
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
                  <td width="200" align="right">�������⣺</td>
                  <td><input name="MesName" type="text" id="MesName" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�������ݣ�</td>
                  <td><textarea name="Content" cols="70" rows="8"></textarea>
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�ƺ���</td>
                  <td><input name="Linkman" type="text" id="Linkman" value="<%=mLinkman%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�Ա�</td>
                  <td><input name="Sex" type="radio" value="����" class="inputnoborder" <%if mSex="����" then response.write ("checked")%> />
                    ����
                    <input type="radio" name="Sex" value="Ůʿ" class="inputnoborder" <%if mSex="Ůʿ" then response.write ("checked")%> />
                    Ůʿ <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">��˾���ƣ�</td>
                  <td><input name="Company" type="text" value="<%=mCompany%>" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">��ϵ��ַ��</td>
                  <td><input name="Address" type="text" value="<%=mAddress%>" size="50" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�������룺</td>
                  <td><input name="ZipCode" type="text" value="<%=mZipCode%>" size="20" maxlength="20" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">��ϵ�绰��</td>
                  <td><input name="Telephone" type="text" id="Telephone" value="<%=mTelephone%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">������룺</td>
                  <td><input name="Fax" type="text" id="Fax" value="<%=mFax%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�ֻ����룺</td>
                  <td><input name="Mobile" type="text" id="Mobile" value="<%=mMobile%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">�������䣺</td>
                  <td><input name="Email" type="text" id="Email" value="<%=mEmail%>" size="30" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">���Ļ���</td>
                  <td><%=SecretFlag%></td>
                </tr>
                <tr>
                  <td align="right">��֤�룺</td>
                  <td><input name="CheckCode" type="text" size="6" maxlength="6" />
                    <a href="javascript:refreshimg()" title="�������������ͼƬ��"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td><input name="Submit" type="submit" value="��������" /></td>
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
  response.write "<a href=""MessageWrite.asp"">ǩд����</a>"
  response.write "<a href=""MessageList.asp"">��������</a>"
  response.write "<a href=""MemberMessage.asp"">�ҵ�����</a>"
end function

function WebLocation()
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""MessageList.asp"" class=""agray"">���Է���</a> - ǩд����"&VbCrLf
end Function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<center>��������</center>"
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