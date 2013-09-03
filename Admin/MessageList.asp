<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|22,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
dim Result,Keyword
Result=request.QueryString("Result")
Keyword=request.QueryString("Keyword")
function PlaceFlag()
  if Result="Search" then
	If Keyword<>"" Then
		Response.Write "留言：列表 -> 检索 -> 关键字：<font color='red'>"&Keyword&"</font>"
	Else
		Response.Write "留言：列表 -> 检索 -> 关键字为空(显示全部留言)"
	End If
  else
    if SortPath<>"" then
      Response.Write "留言：列表 -> <a href='MessageList.asp'>全部</a>"
	  TextPath(SortID)
	else
      Response.Write "留言：列表 -> 全部"
	end if
  end if
end function
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="formSearch" method="post" action="Search.asp?Result=Message">
  <tr>
    <th height="22" sytle="line-height:150%">【留言检索及分类查看】</th>
  </tr>
  <tr>
    <td class="forumRow">关键字：<input name="Keyword" type="text" value="<%=Keyword%>" size="20"> <input name="submitSearch" type="submit" value="搜索留言"></td>
  </tr>
  <tr>
    <td class="forumRow"><%PlaceFlag()%></td>
  </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.asp?Result=Message" method="post" name="formDel">
  <tr>
    <th>ID</th>
	<th>留言人</th>
	<th width="200">标题</th>
	<th>状态</th>
	<th>留言时间</th>
	<th>审核</th>
	<th>回复时间</th>
	<th>操作</th>
	<th>选择</th>
  </tr>
  <% MessageList() %>
  </form>
</table>
<%
function MessageList()
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
      datafrom="Qianbo_Message"
  dim datawhere
      if Result="Search" then
	     datawhere="where MesName like '%" & Keyword &_
		           "%' "
	  else
	    if SortPath<>"" then
		  datawhere="where Instr(SortPath,'"&SortPath&"')>0 "
        else
		  datawhere=""
		end if
	  end if
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
	  Response.Write "<td nowrap class=""forumRow"">"&Guest(rs("MemID"),rs("Linkman"))&"</td>" & vbCrLf
	  if StrLen((rs("MesName")))>41 then
        Response.Write "<td title="&rs("MesName")&" nowrap class=""forumRow"">"&StrLeft(rs("MesName"),39)&"</td>" & vbCrLf
      else
        Response.Write "<td title="&rs("MesName")&" nowrap class=""forumRow"">"&rs("MesName")&"</td>" & vbCrLf
      end if 
      if rs("SecretFlag") then
        Response.Write "<td nowrap class=""forumRow""><font color='red'>悄悄话</font></td>" & vbCrLf
      else
        Response.Write "<td nowrap class=""forumRow"">普通</td>" & vbCrLf
	  end if	  
      Response.Write "<td nowrap class=""forumRow"">"&rs("AddTime")&"</td>" & vbCrLf
      if rs("ViewFlag") then
        Response.Write "<td nowrap align=""center"" class=""forumRow""><font color='blue'>生效</font></td>" & vbCrLf
      else
        Response.Write "<td nowrap align=""center"" class=""forumRow""><font color='red'>未生效</font></td>" & vbCrLf
	  end If
      If rs("ReplyTime") <> "" Then
      ReplyTime = rs("ReplyTime")
	  Else
	  ReplyTime = "<font color=""#CC0000"">暂无回复</font>"
	  End If
      Response.Write "<td nowrap class=""forumRow"">"&ReplyTime&"</td>" & vbCrLf
      Response.Write "<td nowrap align=""center"" class=""forumRow""><a href='MessageEdit.asp?Result=Modify&ID="&rs("ID")&"'>审核、回复</a></td>" & vbCrLf
 	  Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='9' nowrap align=""right"" class=""forumRow""><input type=""submit"" name=""batch"" value=""批量生效"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""批量失效"" onClick=""return test();""> <input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input name='batch' type='submit' value='删除所选' onClick=""return test();""></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td nowrap align='center' colspan='9' class=""forumRow"">暂无留言信息</td></tr>"
  end if
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='9' nowrap class=""forumRow"">" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td class=""forumRow"">共计：<font color='red'>"&idcount&"</font>条记录 页次：<font color='red'>"&page&"</font></strong>/"&pagec&" 每页：<font color='red'>"&pages&"</font>条</td>" & vbCrLf
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
  Response.Write "第<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('请输入需要跳转到的页数并且必须为整数！');this.value='"&Page&"';}"" style='width: 28px;' type='text' value='"&Page&"'>页" & vbCrLf
  Response.Write "<input name='submitSkip' type='button' onClick='GoPage("""&Myself&""")' value='转到'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
end Function

function Guest(ID,Linkman)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    Guest=Linkman
  else
    Guest="<font color='green'>会员：</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"'>"&Linkman&"</a>"
  end if
  rs.close
  set rs=nothing
end function
%>