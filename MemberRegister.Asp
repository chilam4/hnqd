<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>ע���Ա - <% =SiteTitle %></title>
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
                <td width="200" align="right">�û����ƣ�</td>
                <td><input name="MemName" type="text" id="MemName" size="20" maxlength="16" /> <font color="red">*</font>(0-9,a-z,-_������ϵ�3-16���ַ�)</td>
              </tr>
              <tr>
                <td align="right">��ʵ������</td>
                <td><input name="RealName" type="text" id="RealName" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�Ա�</td>
                <td><input name="Sex" type="radio" value="����" checked="checked" class="inputnoborder" />���� <input type="radio" name="Sex" value="Ůʿ" class="inputnoborder" />Ůʿ <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�û����룺</td>
                <td><input name="Password" type="password" id="Password" size="20" maxlength="16" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�ظ��û����룺</td>
                <td><input name="vPassword" type="password" id="vPassword" size="20" maxlength="16" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">���뱣�����⣺</td>
                <td><input name="Question" type="text" id="Question" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">���뱣���𰸣�</td>
                <td><input name="Answer" type="text" id="Answer" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��λ���ƣ�</td>
                <td><input name="Company" type="text" id="Company" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ϸ��ַ��</td>
                <td><input name="Address" type="text" id="Address" size="40" maxlength="100" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�������룺</td>
                <td><input name="ZipCode" type="text" id="ZipCode" size="20" maxlength="20" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ϵ�绰��</td>
                <td><input name="Telephone" type="text" id="Telephone" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">������룺</td>
                <td><input name="Fax" type="text" id="Fax" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�ֻ����룺</td>
                <td><input name="Mobile" type="text" id="Mobile" size="20" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">�������䣺</td>
                <td><input name="Email" type="text" id="Email" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��ַ��</td>
                <td><input name="HomePage" type="text" id="HomePage" size="30" maxlength="50" /> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">��֤�룺</td>
                <td><input name="CheckCode" type="text" size="6" maxlength="6" /> <a href="javascript:refreshimg()" title="�������������ͼƬ��"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input name="Submit" type="submit" value="����ע���Ա����ȡ���²�Ʒ��Ѷ" /></td>
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
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""MemberCenter.asp"" class=""agray"">��Ա����</a> - �ʻ�ע��"&VbCrLf
end function
%>
<script language="javascript">
<!--
function refreshimg(){document.all.checkcode.src="Include/CheckCode/CheckCode.asp";}
//-->
</script>