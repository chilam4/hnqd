<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/Version.asp" -->
<%
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title>��ҵ��վ����ϵͳ</title>
<link href="images/index.css" type="text/css" rel="stylesheet" />
<link href="images/MasterPage.css" type="text/css" rel="stylesheet" />
<link href="images/Guide.css" type="text/css" rel="stylesheet" />
<link href="images/Login.css" type="text/css" rel="stylesheet" />
<script type="text/javascript">
<!--
if(self!=top){
    top.location=self.location;
}
CheckBrowser();
SetFocus();
var closestr=0;
function SetFocus() {
    if(document.Login.UserName.value == "")
    document.Login.UserName.focus();
    else
    document.Login.UserName.select();
}
function CheckForm() {
    if(document.Login.UserName.value == "") {
        alert("�������û�����");
        document.Login.UserName.focus();
        return false;
    }
    if(document.Login.password.value == "") {
        alert("���������룡");
        document.Login.password.focus();
        return false;
    }
    if (document.Login.CheckCode.value == "") {
        alert ("������������֤�룡");
        document.Login.CheckCode.focus();
        return(false);
    }
    if (document.Login.AdminLoginCode.value == "") {
        alert ("���������Ĺ�����֤�룡");
        document.Login.AdminLoginCode.focus();
        return(false);
    }
}
function CheckBrowser() {
    var app=navigator.appName;
    var verStr=navigator.appVersion;
    if(app.indexOf("Netscape") != -1) {
        alert("����������ʾ��\n��ʹ�õ���Netscape��Firefox����������IE����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
    }
     else if(app.indexOf("Microsoft") != -1) {
        if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
        alert("����������ʾ��\n����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
    }
}
function refreshimg(){
    document.all.checkcode.src="../Include/CheckCode/CheckCode.Asp";
}
-->
</script>
</head>

<body id="loginbody">
<form action="CheckLogin.Asp" method="post" name="Login" onSubmit="return CheckForm()">
  <div id="adminboxall">
    <div class="adminboxtop"></div>
    <div id="adminboxmain">
      <div class="menu">
          <input type="image" name="Submit" src="images/admin_menu.gif" style="border-width: 0px; width: 76px; height: 26px;" />
      </div>
    </div>
    <div class="adminboxbottom">
      <div id="login">
        <ul>
          <li class="text">�û�����<br />
            <div class="box1">
              <input name="UserName" type="text" maxlength="20" class="boxcontent" style="font-family: ����;" />
            </div>
          </li>
          <li class="text">�� �룺<br />
            <div class="box2">
              <input name="password" type="password" maxlength="20" class="boxcontent" />
            </div>
          </li>
          
          <li class="text">������֤�룺<br />
            <div class="box3">
              <input name="AdminLoginCode" type="password" maxlength="20" class="boxcontent" />
            </div>
          </li>
          
          <li class="textCode">��֤�룺<br />
            <div class="box4">
              <input name="CheckCode" type="password" maxlength="20" class="boxcontent2" style="ime-mode: disabled;" />
              <a href="javascript:refreshimg()" title="�������������ͼƬ��"><img src="../Include/CheckCode/CheckCode.Asp" style="border: 1px solid #ffffff" align="absmiddle" /></a></div>
          </li>
        </ul>
      </div>
    </div>
    <a href="http://www.qianbo.com.cn/" target="_blank"><img src="images/admin_text.gif" width="186" border="0" height="10" alt="��ҵ��վ����ϵͳ" /></a>
    <div class="clearbox"></div>
  </div>
</form>
</body>
</html>