<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|17,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
dim Result,Keyword,SortID,SortPath
Result=request.QueryString("Result")
Keyword=request.QueryString("Keyword")
SortID=request.QueryString("SortID")
SortPath=request.QueryString("SortPath")
function PlaceFlag()
  if Result="Search" Then
	If Keyword<>"" Then
		Response.Write "招聘：列表 -> 检索 -> 关键字：<font color='red'>"&Keyword&"</font>"
	Else
		Response.Write "招聘：列表 -> 检索 -> 关键字为空(显示全部招聘)"
	End If
  else
    if SortPath<>"" then
      Response.Write "招聘：列表 -> <a href='JobsList.asp'>全部</a>"
	  TextPath(SortID)
	else
      Response.Write "招聘：列表 -> 全部"
	end if
  end if
end function
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="formSearch" method="post" action="Search.asp?Result=Jobs">
  <tr>
    <th height="22" sytle="line-height:150%">【招聘检索及查看】</th>
  </tr>
  <tr>
    <td class="forumRow">关键字：<input name="Keyword" type="text" value="<%=Keyword%>" size="20"> <input name="submitSearch" type="submit" value="搜索招聘"></td>
  </tr>
  <tr>
    <td class="forumRow"><%PlaceFlag()%></td>
  </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.asp?Result=Jobs" method="post" name="formDel">
  <tr>
    <th>ID</th>
	<th>生效</th>
	<th width="200">招聘标题</th>
	<th>职位数量</th>
	<th>结束日期</th>
	<th>发布日期</th>
	<th>应聘人数</th>
	<th>人气</th>
	<th>操作</th>
	<th>选择</th>
  </tr>
  <% JobsList() %>
  </form>
</table>
<%
function JobsList()
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
      datafrom="Qianbo_Jobs"
  dim datawhere
      if Result="Search" then
	     datawhere="where JobName like '%" & Keyword &_
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
      if rs("ViewFlag") then
        Response.Write "<td nowrap align='center' class=""forumRow"" width=""40""><a href=""Conversion.asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=down""><font color='blue'>生效</font></a></td>" & vbCrLf
      else
        Response.Write "<td nowrap align='center' class=""forumRow"" width=""40""><a href=""Conversion.asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=up""><font color='red'>未生效</font></a></td>" & vbCrLf
	  end If
	  if StrLen((rs("JobName")))>40 then
        Response.Write "<td nowrap title='"&rs("JobName")&"' class=""forumRow""><a href='JobsEdit.asp?Result=Modify&ID="&rs("ID")&"''>"&StrLeft(rs("JobName"),37)&"</a></td>" & vbCrLf
      else
        Response.Write "<td nowrap title='"&rs("JobName")&"' class=""forumRow""><a href='JobsEdit.asp?Result=Modify&ID="&rs("ID")&"''>"&rs("JobName")&"</a></td>" & vbCrLf
      end if 
      Response.Write "<td nowrap class=""forumRow"">"&rs("JobNumber")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("EndDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("TalentsNumber")&"人</td>" & vbCrLf
      Response.Write "<td nowrap class=""forumRow"">"&rs("ClickNumber")&"</td>" & vbCrLf
      Response.Write "<td align=""center""nowrap class=""forumRow""><a href='JobsEdit.asp?Result=Add'>添加</a> <a href='JobsEdit.asp?Result=Modify&ID="&rs("ID")&"'>修改</a></td>" & vbCrLf
 	  Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='10' nowrap align=""right"" class=""forumRow""><input type=""submit"" name=""batch"" value=""批量生效"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""批量失效"" onClick=""return test();""> <input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input name='batch' type='submit' value='删除所选' onClick=""return test();""></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td nowrap align='center' colspan='10' class=""forumRow"">暂无企业招聘</td></tr>"
  end if
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='10' nowrap class=""forumRow"">" & vbCrLf
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
%>