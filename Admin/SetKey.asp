<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<%
if Instr(session("AdminPurview"),"|36,")=0 then
  response.write "<center>��û�й����ģ���Ȩ�ޣ�</center>"
  response.end
End If
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.asp?Result=SiteLink" method="post" name="formDel">
  <tr>
    <th width="8">ID</th>
	<th align="left">��������</th>
	<th align="left">���ӵ�ַ</th>
	<th width="60" align="left">���ȼ���</th>
	<th width="60" align="left">�滻����</th>
	<th width="60" align="left">�򿪷�ʽ</th>
	<th width="30">״̬</th>
	<th align="left" width="60">����</th>
	<th width="28">����</th>
  </tr>
  <% SetLinkList() %>
  </form>
</table>
<%
function SetLinkList()
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
      datafrom="Qianbo_Link"
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
      Response.Write "<td nowrap class=""leftrow"">"&rs("ID")&"</td>" & vbCrLf
	  Response.Write "<td nowrap class=""leftrow""><a href=""LinkEdit.asp?id="&rs("ID")&"&Result=Modify"">"&rs("Text")&"</a></td>" & vbCrLf
      Response.Write "<td nowrap class=""leftrow"">"&rs("Link")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""leftrow"">"&rs("Order")&"</td>" & vbCrLf
	  Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
	  If rs("Replace") = 0 Then
	  Response.Write "����"
	  Else
	  Response.Write rs("Replace")
	  End If
	  Response.Write "</td>" & vbCrLf
	  Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
	  Select Case Cstr(LCase(rs("Target")))
	  Case "0"
		Response.Write "<font color=""blue"">ԭ����</font>"
	  Case "1"
		Response.Write "�´���"
	  Case Else
		Response.Write rs("Target")
	  End Select
	  Response.Write "</td>" & vbCrLf
	  Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
	  Select Case Cstr(LCase(rs("State")))
	  Case "0"
		Response.Write "<font color=""red"">����</font>"
	  Case "1"
		Response.Write "����"
	  Case Else
		Response.Write rs("State")
	  End Select
	  Response.Write "</td>" & vbCrLf
      Response.Write "<td nowrap class=""leftrow""><a href=""LinkEdit.asp?Result=Add"">���</a> <a href=""LinkEdit.asp?id="&rs("ID")&"&Result=Modify"">�޸�</a></td>" & vbCrLf
      Response.Write "<td nowrap class=""centerrow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='9' nowrap class=""forumRow"" align=""right""><input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""ȫѡ""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""��ѡ""> <input type=""submit"" name=""batch"" value=""������Ч"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""����ʧЧ"" onClick=""return test();""> <input name='batch' type='submit' value='ɾ����ѡ' onClick=""return test();""></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td colspan='9' nowrap class=""centerrow"">�����������</td></tr>"
  end if
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='9' nowrap class=""leftrow"">" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td class=""leftrow"">��<font color='red'> "&idcount&" </font>����¼ ҳ�Σ�<font color='red'>"&page&"</font></strong>/"&pagec&"&nbsp;ÿҳ��<font color='red'>"&pages&"</font>��</td>" & vbCrLf
  Response.Write "<td class=""forumRow"" align=""right"">" & vbCrLf
  pagenmin=page-pagenc
  pagenmax=page+pagenc
  if(pagenmin<1) then pagenmin=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='font-size: 14px; font-family: webdings'>9</font></a> ")
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='font-size: 14px; font-family: webdings'>7</font></a> ")
  if(pagenmax>pagec) then pagenmax=pagec
  for i = pagenmin to pagenmax
	if(i=page) then
	  response.write (" <font color='red'>"& i &"</font> ")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write (" <a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='font-size: 14px; font-family: webdings'>8</font></a> ")
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='font-size: 14px; font-family: webdings'>:</font></a> ")
  Response.Write "��ת������ <input name='SkipPage' style=""width: 30px"" onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('������ת���������֣�');this.value='"&Page&"';}"" type='text' value='"&Page&"'> ҳ" & vbCrLf
  Response.Write "<input name='submitSkip' type='button' onClick='GoPage("""&Myself&""")' value='ת��'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
end function
%>