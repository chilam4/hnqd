<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|38,")=0 then
	response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
	response.end
end If
dim act,tablename
act=trim(request.QueryString("act"))
tablename=trim(request.Form("tablename"))
if(act="create") then
	conn.execute("Create table "&tablename&"(id AUTOINCREMENT(1,1),primary key(id))")
	response.Redirect("SQL.asp")
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<style type="text/css">
<!--
.tb {
	float:left;
	margin:0;
	padding:0;
	text-align:center;
	width:128px;
}
-->
</style>
<br />
<table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" sytle="line-height:150%">�����ݿ�����</th>
  </tr>
  <tr>
    <td class="forumRow" style="line-height:300%"><%
dim i,str
i=1
set rsSchema = conn.openSchema(20)
rsSchema.movefirst
Do Until rsSchema.EOF
if rsSchema("TABLE_TYPE")="TABLE" then
	response.Write "<div class=""tb""><a href=""SQLEdit.asp?tablename="&rsSchema("table_name")&"""><img src=""images/table.gif"" border=""0"" align=""absmiddle""><br />"&rsSchema("TABLE_NAME")&"</a></div>"
end if
rsSchema.movenext
i=i+1
Loop
%>
    </td>
  </tr>
  <form name="form1" method="post" action="?act=create">
    <tr>
      <td class="centerrow">���ݿ�����ƣ�
        <input type="text" name="tablename">
        <input type="submit" name="Submit" value="�����±�"></td>
    </tr>
  </form>
</table>
<br />
<center>
  <font color="red">���棺�����ܽ�ֱ�Ӳ����������ݿ�ṹ���������רҵ��Աʹ�á��ڽ����κβ���ǰ�뱸���������ݣ�</font>
</center>
<br />