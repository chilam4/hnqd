<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<%
if Instr(session("AdminPurview"),"|39,")=0 then
	response.write "<center>��û�й����ģ���Ȩ�ޣ�</center>"
	response.end
end If
dim i,tablename
tablename=trim(request.QueryString("tablename"))
if(len(tablename)<1) then
	response.write "<script language='JavaScript'>alert('���ݱ��������');" & "history.back()" & "</script>"
	response.End()
end If
dim Action,rsCheckAdd,rs,sql
Action=request.QueryString("Action")
if Action="SaveEdit" then
	fieldname=trim(request.Form("fieldname"))
	if(len(fieldname)<1) then
 		response.write "<script language='JavaScript'>alert('�������ֶ�����');" & "history.back()" & "</script>"
 		response.End()
	end if
	fieldtype=trim(request.Form("fieldtype"))
	if(len(fieldtype)<1) then
 		response.write "<script language='JavaScript'>alert('��ѡ���ֶ����ͣ�');" & "history.back()" & "</script>"
 		response.End()
	end if
	if fieldtype="varchar" then
		charlen=Cint(request.Form("varchar_len"))
		addsql="alter table "&tablename&" add "&fieldname&" "&fieldtype&" ("&charlen&") null"
	else
		addsql="alter table "&tablename&" add "&fieldname&" "&fieldtype
	end if
	conn.execute(addsql)
	Response.Write "<script language=javascript>alert('���ݱ� "&tablename&" ���ֶ� "&fieldname&" ��ӳɹ���');window.location.href='"&request.servervariables("http_referer")&"';</script>"
end if
set rs=server.createobject("adodb.recordset")
rs.open "select top 1 * from "&tablename,conn,3,1
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="22" sytle="line-height:150%">
    <th align="left" width="200" style="padding-left: 8px">�ֶ�����</th>
    <th align="left" style="padding-left: 8px">�ֶ�����</th>
  </tr>
  <%For i=0 To rs.fields.count-1%>
  <tr>
    <td class="forumRow" style="padding-left: 8px"><%=rs(i).name%></td>
    <td class="forumRow" style="padding-left: 8px"><%
if rs.fields(i).type="3" then
	response.write "������"
if rs.fields(i).Attributes="16" then response.write " �Զ�����ֶ�"
elseif rs.fields(i).type="202" then
	response.write "�ı���"
	response.write "����"&rs.fields(i).DefinedSize
elseif rs.fields(i).type="2" then
	response.write "����"
elseif rs.fields(i).type="11" then
	response.write "��/��"
elseif rs.fields(i).type="135" Or rs.fields(i).type="7" then
	response.write "����/ʱ��"
elseif rs.fields(i).type="203" then
	response.write "��ע"
elseif rs.fields(i).type="6" then
	response.write "����"
elseif rs.fields(i).type="205" then
	response.write "OLE ����"
else
	response.write "δ֪"&rs.fields(i).type
end if
%></td>
  </tr>
  <%Next%>
</table>
<%
rs.close
set rs=nothing
%>
<br />
<script language="javascript">
function seleChan(str){
	if(str=="varchar"){
		document.getElementById("charlen").style.display="";
	}
	else{
		document.getElementById("charlen").style.display="none";
	}
}
</script>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="form1" method="post" action="?Action=SaveEdit&tablename=<%=tablename%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">��������ֶΡ�</th>
    </tr>
    <tr>
      <td width="200" class="forumRow" align="right">�ֶ����ƣ�</td>
      <td class="forumRowHighlight"><input name="fieldname" type="text" style="width: 150px">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">�ֶ����ͣ�</td>
      <td class="forumRowHighlight"><select name="fieldtype" style="width: 100px" onChange="seleChan(this.options[this.selectedIndex].value)">
          <option value="int">������</option>
          <option value="smallint">����</option>
          <option value="varchar">�ı�</option>
          <option value="datetime">����/ʱ��</option>
          <option value="memo">��ע</option>
          <option value="money">����</option>
          <option value="bit">��/��</option>
        </select>
        <font color="red">*</font></td>
    </tr>
    <tr id="charlen" style="display:none; float:inherit;">
      <td class="forumRow" align="right">�ֶγ��ȣ�</td>
      <td class="forumRowHighlight"><input name="varchar_len" type="text" id="varchar_len" style="width: 100px"></td>
    </tr>
    <tr>
      <td class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="��������">
        <input type="button" value="������һҳ" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<br />
<center>
  <font color="red">���棺�����ܽ�ֱ�Ӳ����������ݿ�ṹ���������רҵ��Աʹ�á��ڽ����κβ���ǰ�뱸���������ݣ�</font>
</center>
<br />