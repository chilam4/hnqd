<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Admin_htmlconfig.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<script language="javascript">
<!--
function setfile(){
    var arr = showModalDialog("eWebEditor/customDialog/file.htm", "", "dialogWidth:22em; dialogHeight:9em; status:0;help=no");
    if (arr ==null){
        alert("ϵͳ��ʾ����ǰû���ϴ��ļ���");
    }
    if (arr !=null){
        editForm.FileUrl.value=arr;
    }
}
//-->
</script>
<%
if Instr(session("AdminPurview"),"|16,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
  response.end
end if
dim Result
Result=request.QueryString("Result")
dim ID,DownName,ViewFlag,SortName,SortID,SortPath
dim FileSize,FileUrl,CommendFlag,GroupID,GroupIdName,Exclusive,Content,SeoKeywords,SeoDescription
ID=request.QueryString("ID")
call DownEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="DownEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">��<%If Result = "Add" then%>���<%ElseIf Result = "Modify" then%>�޸�<%End If%>���ء�</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">���ر��⣺</td>
      <td width="80%" class="forumRowHighlight"><input name="DownName" type="text" id="DownName" style="width: 280" value="<%=DownName%>" maxlength="100">
        �Ƿ���Ч��<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>
		�Ƿ��Ƽ���<input name="CommendFlag" type="checkbox" value="1" <%if CommendFlag then response.write ("checked")%>> <font color="red">*</font> <input type="button" name="btn" value="���Ʊ���" title="���Ʊ��⵽��MetaDescription��MetaKeywords" onclick="CopyWebTitle(document.editForm.DownName.value);"></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">MetaKeywords��</td>
      <td width="80%" class="forumRowHighlight"><input name="SeoKeywords" type="text" id="SeoKeywords" style="width: 500" value="<%=SeoKeywords%>" maxlength="250"></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">MetaDescription��</td>
      <td width="80%" class="forumRowHighlight"><input name="SeoDescription" type="text" id="SeoDescription" style="width: 500" value="<%=SeoDescription%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�������</td>
      <td class="forumRowHighlight"><input name="SortID" type="text" id="SortID" style="width: 18; background-color:#fffff0" value="<%=SortID%>" readonly> <input name="SortPath" type="text" id="SortPath" style="width: 70; background-color:#fffff0" value="<%=SortPath%>" readonly> <input name="SortName" type="text" id="SortName" value="<%=SortName%>" style="width: 180; background-color:#fffff0" readonly> <a href="javaScript:OpenScript('SelectSort.asp?Result=Download',500,500,'')">ѡ�����</a> <font color="red">*</font></td>
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
      <td width="20%" align="right" class="forumRow">���ص�ַ��</td>
      <td width="80%" class="forumRowHighlight"><input name="FileUrl" type="text" id="FileUrl" style="width: 280" value="<%=FileUrl%>" maxlength="100"> <input type="button" value="�ϴ��ļ�" onClick="setfile();"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">�ļ���С��</td>
      <td width="80%" class="forumRowHighlight"><input name="FileSize" type="text" id="FileSize" style="width: 280" value="<%=FileSize%>" maxlength="100"> MB <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ϸ˵����</td>
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
sub DownEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("DownName")))<3 then
      response.write ("<script language='javascript'>alert('����д���ر��⣡');history.back(-1);</script>")
      response.end
    end If
    if Request.Form("SortID")="" and Request.Form("SortPath")="" then
      response.write ("<script language='javascript'>alert('��ѡ���������࣡');history.back(-1);</script>")
      response.end
    end If
    if Request.Form("FileUrl")="" then
      response.write ("<script language='javascript'>alert('����д���ص�ַ��');history.back(-1);</script>")
      response.end
    end If
    if Request.Form("FileSize")="" then
      response.write ("<script language='javascript'>alert('����д�����ļ���С��');history.back(-1);</script>")
      response.end
    end If
    if Request.Form("Content")="" then
      response.write ("<script language='javascript'>alert('����д��ϸ˵����');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then
	  sql="select * from Qianbo_Download"
      rs.open sql,conn,1,3
      rs.addnew
      rs("DownName")=trim(Request.Form("DownName"))
	  if Request.Form("ViewFlag")=1 then
	  rs("ViewFlag")=Request.Form("ViewFlag")
	  else
	  rs("ViewFlag")=0
	  end if
	  rs("SortID")=Request.Form("SortID")
	  rs("SortPath")=Request.Form("SortPath")
	  if Request.Form("CommendFlag")=1 then
	  rs("CommendFlag")=Request.Form("CommendFlag")
	  else
	  rs("CommendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("FileSize")=trim(Request.Form("FileSize"))
	  rs("FileUrl")=trim(Request.Form("FileUrl"))
	  rs("Content")=rtrim(Request.Form("Content"))
	  rs("SeoKeywords")=trim(Request.Form("SeoKeywords"))
	  rs("SeoDescription")=trim(Request.Form("SeoDescription"))
	  rs("AddTime")=now()
	  rs("UpdateTime")=now()
	  rs.update
	  rs.close
	  set rs=Nothing
	  set rs=server.createobject("adodb.recordset")
	  sql="select top 1 ID from Qianbo_Download order by ID desc"
	  rs.open sql,conn,1,1
	  ID=rs("ID")
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"","DownView.asp","ID=",ID,"","")
	  End If
	end if
	if Result="Modify" then
      sql="select * from Qianbo_Download where ID="&ID
      rs.open sql,conn,1,3
      rs("DownName")=trim(Request.Form("DownName"))
	  if Request.Form("ViewFlag")=1 then
	  rs("ViewFlag")=Request.Form("ViewFlag")
	  else
	  rs("ViewFlag")=0
	  end if
	  rs("SortID")=Request.Form("SortID")
	  rs("SortPath")=Request.Form("SortPath")
	  if Request.Form("CommendFlag")=1 then
	  rs("CommendFlag")=Request.Form("CommendFlag")
	  else
	  rs("CommendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("FileSize")=trim(Request.Form("FileSize"))
	  rs("FileUrl")=trim(Request.Form("FileUrl"))
	  rs("Content")=rtrim(Request.Form("Content"))
	  rs("SeoKeywords")=trim(Request.Form("SeoKeywords"))
	  rs("SeoDescription")=trim(Request.Form("SeoDescription"))
	  rs("UpdateTime")=now()
	  rs.update
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"","DownView.asp","ID=",ID,"","")
	  End If
	end if
    if ISHTML = 1 then
    response.write "<script language='javascript'>alert('���óɹ�����ؾ�̬ҳ���Ѹ��£�');location.replace('DownList.asp');</script>"
	Else
	response.write "<script language='javascript'>alert('���óɹ���');location.replace('DownList.asp');</script>"
	End If
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Download where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("<center>���ݿ��¼��ȡ����</center>")
        response.end
      end if
	  DownName=rs("DownName")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  CommendFlag=rs("CommendFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
	  FileSize=rs("FileSize")
	  FileUrl=rs("FileUrl")
      Content=rs("Content")
	  SeoKeywords=rs("SeoKeywords")
	  SeoDescription=rs("SeoDescription")
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

Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_DownSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>