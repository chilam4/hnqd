<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|32,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й�����ģ���Ȩ�ޣ�</font></div>")
  response.end
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
dim Result
Result=request.QueryString("Result")
dim ID,GroupID,GroupName,GroupLevel,Explain,AddTime,RanNum
ID=request.QueryString("ID")
randomize timer
RanNum=Int((8999)*Rnd +1009)
if ID="" then GroupID=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum
if Result<>"" then
  call MemGroupEdit()
end if
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.asp?Result=MemGroup" method="post" name="formDel">
  <tr>
    <th>ID</th>
	<th>����</th>
	<th>�������</th>
	<th>Ȩ��ֵ</th>
	<th>˵��</th>
	<th>����ʱ��</th>
	<th>����</th>
	<th>ѡ��</th>
  </tr>
  <% MemGroupList() %>
  </form>
</table>
<%
sub MemGroupEdit()
  dim Action,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then
	  sql="select * from Qianbo_MemGroup"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language='javascript'>alert('����д��Ա�������(6-16���ַ���3-8������)��');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupID")=Request.Form("GroupID")
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("GroupLevel")=trim(Request.Form("GroupLevel"))
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if
	if Result="Modify" then
      sql="select * from Qianbo_MemGroup where ID="&ID
      rs.open sql,conn,1,3
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language='javascript'>alert('����д��Ա�������(6-16���ַ���3-8������)��');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("GroupLevel")=trim(Request.Form("GroupLevel"))
	  rs("Explain")=trim(Request.Form("Explain"))
      conn.execute("Update Qianbo_Members set GroupName='"&trim(Request.Form("GroupName"))&"' where GroupID='"&trim(Request.Form("GroupID"))&"'")
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('���óɹ���');location.replace('MemGroup.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_MemGroup where ID="& ID
      rs.open sql,conn,1,1
	  if rs.RecordCount=0 then
        response.write "<script language='javascript'>alert('�޴˼�¼��');history.back(-1)</script>"
        response.end
	  end if
	  ID=rs("ID")
      GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  GroupLevel=rs("GroupLevel")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing
	end if
  end if
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editMemGroup" method="post" action="MemGroup.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>"">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">���޸Ļ�Ա���</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">ID��</td>
      <td width="80%" class="forumRowHighlight"><input name="ID" type="text" id="ID" style="width: 80" value="<%if ID="" then response.write ("�Զ�") else response.write (ID) end if%>" maxlength="6" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���ţ�</td>
      <td class="forumRowHighlight"><input name="GroupID" type="text" id="GroupID" style="width: 180" value="<%=GroupID%>" maxlength="18" readonly> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������ƣ�</td>
      <td class="forumRowHighlight"><input name="GroupName" type="text" id="GroupName" style="width: 180" value="<%=GroupName%>"> <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">Ȩ��ֵ��</td>
      <td class="forumRowHighlight"><input name="GroupLevel" type="text" id="GroupLevel" style="width: 80" value="<%=GroupLevel%>" onChange="if(/\D/.test(this.value)){alert('����Ȩ��ֵ������������');}"> <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��ע��</td>
      <td class="forumRowHighlight"><textarea name="Explain" rows="8" id="Explain" style="width: 500"><%=Explain%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="����"></td>
    </tr>
  </form>
</table>
<%
End Sub

function MemGroupList()
  dim idCount
  dim pages
      pages=20
  dim pagec
  dim page
      page=clng(request("Page"))
  dim pagenc
      pagenc=2
  dim pagenmax
  dim pagenmin
  dim datafrom
      datafrom="Qianbo_MemGroup"
  dim datawhere
      datawhere=""
  dim sqlid
  dim Myself,PATH_INFO,QUERY_STRING
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis
      taxis="order by id desc"
  dim i
  dim rs,sql
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")
  if(idcount>0) then
    if(idcount mod pages=0)then
	  pagec=int(idcount/pages)
   	else
      pagec=int(idcount/pages)+1
    end if
    sql="select id from ["& datafrom &"] " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("id")
	  else
	    sqlid=sqlid &","&rs("id")
	  end if
	  rs.movenext
    next
  end if
  if(idcount>0 and sqlid<>"") then
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)
	  Response.Write "<tr>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("GroupID")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("GroupName")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow""><font color='blue'>"&rs("GroupLevel")&"</font></td>" & vbCrLf
	  if len(rs("Explain"))>24 then
        Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&left(rs("Explain"),22)&"��</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&rs("Explain")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap class=""forumRow"">"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td align=""center""nowrap class=""forumRow""><a href='MemGroup.asp?Result=Modify&ID="&rs("ID")&"'>�޸�</a></td>" & vbCrLf
	  if rs("ID")=1 or rs("ID")=2 Then
	  Response.Write "<td nowrap align='center' class=""forumRow""><font color=""red"">��</font></td>" & vbCrLf
	  Else
 	  Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("GroupID")&"'></td>" & vbCrLf
	  End If
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='10' nowrap align=""right"" class=""forumRow""><input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""ȫѡ""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""��ѡ""> <input name='submitDelSelect' type='button' id='submitDelSelect' value='ɾ����ѡ' onClick='ConfirmDel(""�Ƿ�ȷ��ɾ����ɾ�����ָܻ���"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td nowrap align='center' colspan='10' class=""forumRow"">���޻�Ա���</td></tr>"
  end if
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='10' nowrap class=""forumRow"">" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td class=""forumRow"">���ƣ�<font color='red'>"&idcount&"</font>����¼ ҳ�Σ�<font color='red'>"&page&"</font></strong>/"&pagec&" ÿҳ��<font color='red'>"&pages&"</font>��</td>" & vbCrLf
  Response.Write "<td align='right'>" & vbCrLf
  pagenmin=page-pagenc
  pagenmax=page+pagenc
  if(pagenmin<1) then pagenmin=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='font-size: 14px; font-family: Webdings'>9</font></a> ")
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='font-size: 14px; font-family: Webdings'>7</font></a> ")
  if(pagenmax>pagec) then pagenmax=pagec
  for i = pagenmin to pagenmax
	if(i=page) then
	  response.write (" <font color='red'>"& i &"</font> ")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write (" <a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='font-size: 14px; font-family: Webdings'>8</font></a> ")
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='font-size: 14px; font-family: Webdings'>:</font></a> ")
  Response.Write "��<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('��������Ҫ��ת����ҳ�����ұ���Ϊ������');this.value='"&Page&"';}"" style='width: 28px;' type='text' value='"&Page&"'>ҳ" & vbCrLf
  Response.Write "<input name='submitSkip' type='button' onClick='GoPage("""&Myself&""")' value='ת��'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
end Function
%>