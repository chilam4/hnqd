<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<%
if Instr(session("AdminPurview"),"|37,")=0 then
  response.write "<center>��û�й����ģ���Ȩ�ޣ�</center>"
  response.end
end If
dim Result
Result=request.QueryString("Result")
dim ID,mText,mDescription,mLink,mOrder,mReplace,mTarget,mState
ID=request.QueryString("ID")
call LinkEdit()
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="LinkEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">��<%If Result = "Add" then%>���<%ElseIf Result = "Modify" then%>�޸�<%End If%>վ�����ӡ�</th>
    </tr>
    <tr>
      <td width="200" class="forumRow" align="right">�������֣�</td>
      <td class="forumRowHighlight"><input name="mText" type="text" id="mText" style="width: 300" value="<%=mText%>"> ��Ҫ�滻���ı� <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">����������</td>
      <td class="forumRowHighlight"><input name="mDescription" type="text" id="mDescription" style="width: 500" value="<%=mDescription%>"><br />
	  ���ӵ��������ݣ�������ʹ���������֣����ѡ����������|�ָ���ͬ��������</td>
    </tr>
	<tr>
      <td class="forumRow" align="right">���ӵ�ַ��</td>
      <td class="forumRowHighlight"><input name="mLink" type="text" id="mLink" style="width: 500" value="<%=mLink%>" maxlength="100"> �������ֶ�Ӧ�����ӵ�ַ <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">���ȼ���</td>
      <td class="forumRowHighlight"><input name="mOrder" type="text" id="mOrder" style="width: 300" value="<%If mOrder <> "" Then Response.Write ""&mOrder&"" Else Response.Write "1" End If%>"> ����Խ������ȨԽ�� <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">�滻������</td>
      <td class="forumRowHighlight"><input name="mReplace" type="text" id="mReplace" style="width: 300" value="<%If mReplace <> "" Then Response.Write ""&mReplace&"" Else Response.Write "1" End If%>"> �滻���ӵĴ���(Ϊ0ʱ���滻ȫ��) <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">�򿪷�ʽ��</td>
      <td class="forumRowHighlight"><input <%If mTarget = 0 Then Response.Write "checked='checked'"%> name="mTarget" type="radio" value="0" />ԭ���� <input <%If mTarget = "" Or mTarget = 1 Then Response.Write "checked='checked'"%> name="mTarget" type="radio" value="1" />�´��� <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">״̬��</td>
      <td class="forumRowHighlight"><input <%If mState = 0 Then Response.Write "checked='checked'"%> name="mState" type="radio" value="0" />���� <input <%If mState = "" Or mState = 1 Then Response.Write "checked='checked'"%> name="mState" type="radio" value="1" />���� <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="��������">
        <input type="button" value="������һҳ" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<br />
<%
sub LinkEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("mText")))<1 or len(trim(request.Form("mLink")))<1 or len(trim(request.Form("mOrder")))<1 or len(trim(request.Form("mReplace")))<1 then
      response.write ("<script language='javascript'>alert('��������дվ�����Ӹ������ݣ�');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then
	  sql="select * from Qianbo_Link"
      rs.open sql,conn,1,3
      rs.addnew
	  rs("Text")=trim(Request.Form("mText"))
	  rs("Description")=trim(Request.Form("mDescription"))
	  rs("Link")=trim(Request.Form("mLink"))
	  rs("Order")=trim(Request.Form("mOrder"))
	  rs("Replace")=trim(Request.Form("mReplace"))
	  rs("Target")=trim(Request.Form("mTarget"))
	  rs("State")=trim(Request.Form("mState"))
	end if
	if Result="Modify" then
      sql="select * from Qianbo_Link where ID="&ID
      rs.open sql,conn,1,3
	  rs("Text")=trim(Request.Form("mText"))
	  rs("Description")=trim(Request.Form("mDescription"))
	  rs("Link")=trim(Request.Form("mLink"))
	  rs("Order")=trim(Request.Form("mOrder"))
	  rs("Replace")=trim(Request.Form("mReplace"))
	  rs("Target")=trim(Request.Form("mTarget"))
	  rs("State")=trim(Request.Form("mState"))
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('���óɹ���');location.replace('SetKey.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Link where ID="& ID
      rs.open sql,conn,1,1
	  mText=rs("Text")
	  mDescription=rs("Description")
	  mLink=rs("Link")
	  mOrder=rs("Order")
	  mReplace=rs("Replace")
	  mTarget=rs("Target")
	  mState=rs("State")
	  rs.close
      set rs=nothing
	end if
  end if
end sub
%>