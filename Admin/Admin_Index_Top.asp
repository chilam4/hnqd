<!--#include file="CheckAdmin.asp"-->
<!--#include file="../Include/Version.asp" -->
<html>
<head>
<title>�����������˵�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
a:link {
	color:#ffffff;
	text-decoration:none
}
a:hover {
	color:#ffffff;
}
a:visited {
	color:#f0f0f0;
	text-decoration:none
}
.spa {
	font-size: 9pt;
	filter: glow(color=#0f42a6, strength=2) dropshadow(color=#0f42a6, offx=2, offy=1, );
	color: #8aade9;
	font-family: "����";
	padding-right: 8px;
}
img {
filter:alpha(opacity:100);
chroma(color=#ffffff)
}
</style>
<base target="main">
<script language="JavaScript" type="text/JavaScript">
function preloadImg(src) {
  var img=new Image();
  img.src=src
}
preloadImg("Images/admin_top_open.gif");

var displayBar=true;
function switchBar(obj) {
  if (displayBar) {
    parent.frame.cols="0,*";
    displayBar=false;
    obj.src="Images/admin_top_open.gif";
    obj.title="����߹������˵�";
  } else {
    parent.frame.cols="200,*";
    displayBar=true;
    obj.src="Images/admin_top_close.gif";
    obj.title="�ر���߹������˵�";
  }
}
</script>
</head>

<body background="Images/admin_top_bg.gif" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr valign="middle">
    <td width="60"><img onClick="switchBar(this)" src="Images/admin_top_close.gif" title="�ر���߹������˵�" style="cursor:hand"></td>
    <td align="right" class="spa">�汾�ţ���վ����ϵͳ <%=Str_Soft_Version%></td>
  </tr>
</table>
</html>