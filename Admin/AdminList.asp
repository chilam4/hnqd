<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|28,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
  response.end
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.asp?Result=Administrators" method="post" name="formDel">
  <tr>
    <th>ID</th>
	<th>��Ч</th>
	<th>�û���</th>
	<th>����Ȩ��</th>
	<th>����Ȩ��</th>
	<th>����¼ʱ��</th>
	<th>����¼IP</th>
	<th>����ʱ��</th>
	<th>����</th>
	<th>ѡ��</th>
  </tr>
  <% AdminList() %>
  </form>
</table>
<%
function AdminList()
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
      datafrom="Qianbo_Admin"
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
    sql="select [ID],[Working],[AdminName],[UserName],[AdminPurview],[Explain],[LastLoginTime],[LastLoginIP],[AddTime]from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)
	  Response.Write "<tr>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("ID")&"</td>" & vbCrLf
      if rs("working") then
        Response.Write "<td nowrap align='center' class=""forumRow""><font color='blue'>��Ч</font></td>" & vbCrLf
      else
        Response.Write "<td nowrap align='center' class=""forumRow""><font color='red'>δ��Ч</font></td>" & vbCrLf
	  end If
      Response.Write "<td nowrap class=""forumRow"">"&rs("AdminName")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("UserName")&"</td>" & vbCrLf
	  if len(rs("AdminPurview"))>28 then
        Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&left(rs("AdminPurview"),25)&"��</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&rs("AdminPurview")&"</td>" & vbCrLf
      end if 	
      Response.Write "<td nowrap class=""forumRow"">"&rs("LastLoginTime")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("LastLoginIP")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td align=""center""nowrap class=""forumRow""><a href='AdminEdit.asp?Result=Modify&ID="&rs("ID")&"'>�޸�</a></td>" & vbCrLf
	  if rs("ID")=1 Then
	  Response.Write "<td nowrap align='center' class=""forumRow""><font color='red'>��</font></td>" & vbCrLf
	  Else
 	  Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
	  End If
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='10' nowrap align=""right"" class=""forumRow""><input type=""submit"" name=""batch"" value=""������Ч"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""����ʧЧ"" onClick=""return test();""> <input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""ȫѡ""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""��ѡ""> <input name='batch' type='submit' value='ɾ����ѡ' onClick=""return test();""></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td nowrap align='center' colspan='10' class=""forumRow"">������վ����Ա</td></tr>"
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