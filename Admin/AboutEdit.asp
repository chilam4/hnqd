<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Admin_htmlconfig.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
if Instr(session("AdminPurview"),"|10,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
  response.end
end if
dim Result
Result=request.QueryString("Result")
dim ID,AboutName,ViewFlag,Content
dim GroupID,GroupIdName,Exclusive,ChildFlag
ID=request.QueryString("ID")
call AboutEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editAboutForm" method="post" action="AboutEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">��<%If Result = "Add" then%>���<%ElseIf Result = "Modify" then%>�޸�<%End If%>��ҵ��Ϣ��</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">��Ϣ���⣺</td>
      <td width="80%" class="forumRowHighlight"><input name="AboutName" type="text" id="AboutName" style="width: 280" value="<%=AboutName%>" maxlength="100">
        �Ƿ���Ч��<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>
        �Ƿ��ҳ��<input name="ChildFlag" type="checkbox" value="1" <%if ChildFlag then response.write ("checked")%>> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�Ķ�Ȩ�ޣ�</td>
      <td class="forumRowHighlight"><select name="GroupID">
          <% call SelectGroup() %>
        </select>
        <input name="Exclusive" type="radio" value="&gt;=" <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>>
        ����
        <input type="radio" <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
        ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ϣ���ݣ�</td>
      <td class="forumRowHighlight"><textarea name="Content" id="Content" style="display:none"><%=Server.HTMLEncode((Content))%></textarea>
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=Content&style=coolblue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="����"> <input type="button" value="������һҳ" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
sub AboutEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("AboutName")))<1 then
      response.write ("<script language='javascript'>alert('����д��Ϣ���⣡');history.back(-1);</script>")
      response.end
    end If
    if trim(request.Form("Content"))="" then
      response.write ("<script language='javascript'>alert('����д��Ϣ���ݣ�');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then
	  sql="select * from Qianbo_About"
      rs.open sql,conn,1,3
      rs.addnew
      rs("AboutName")=trim(Request.Form("AboutName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=rtrim(Request.Form("Content"))
    GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  if Request.Form("ChildFlag")=1 then
      rs("ChildFlag")=Request.Form("ChildFlag")
	    rs("Sequence")=999
	  else
      rs("ChildFlag")=0
	    rs("Sequence")=99
	  end if
	  rs("AddTime")=now()
	  rs("UpdateTime")=now()
	  rs.update
	  rs.close
	  set rs=Nothing
	  set rs=server.createobject("adodb.recordset")
	  sql="select top 1 ID from Qianbo_About order by ID desc"
	  rs.open sql,conn,1,1
	  ID=rs("ID")
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&AboutNameDiy&""&Separated&""&ID&"."&HTMLName&"","About.asp","ID=",ID,"","")
	  End If
	end if
	if Result="Modify" then
      sql="select * from Qianbo_About where ID="&ID
      rs.open sql,conn,1,3
      rs("AboutName")=trim(Request.Form("AboutName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
	  GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  if Request.Form("ChildFlag")=1 then
      rs("ChildFlag")=Request.Form("ChildFlag")
	    rs("Sequence")=999
	  else
      rs("ChildFlag")=0
	  end if
	  rs("UpdateTime")=now()
	  rs.update
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&AboutNameDiy&""&Separated&""&ID&"."&HTMLName&"","About.asp","ID=",ID,"","")
	  End If
	end if
    if ISHTML = 1 then
    response.write "<script language='javascript'>alert('���óɹ�����ؾ�̬ҳ���Ѹ��£�');location.replace('AboutList.asp');</script>"
	Else
	response.write "<script language='javascript'>alert('���óɹ���');location.replace('AboutList.asp');</script>"
	End If
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_About where ID="& ID
      rs.open sql,conn,1,1
	  AboutName=rs("AboutName")
	  ViewFlag=rs("ViewFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
      Content=rs("Content")
	  ChildFlag=rs("ChildFlag")
	  rs.close
      set rs=nothing
	end if
  end if
end sub

sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from Qianbo_MemGroup"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("δ�����")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"���橾"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>