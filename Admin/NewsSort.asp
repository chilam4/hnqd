<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|6,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
Dim Action
Action=request.QueryString("Action")
Select Case Action
  Case "Add"
	addFolder
  	CallFolderView()
  Case "Del"
    Dim rs,sql,SortPath
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From Qianbo_NewsSort where ID="&request.QueryString("id")
    rs.open sql,conn,1,1
	SortPath=rs("SortPath")
	conn.execute("delete from  Qianbo_NewsSort where Instr(SortPath,'"&SortPath&"')>0")
    conn.execute("delete from  Qianbo_News where Instr(SortPath,'"&SortPath&"')>0")
    response.write ("<script language='javascript'>alert('成功删除本类、子类及所有下属信息条目！');location.replace('NewsSort.asp');</script>")
  Case "Save"
	saveFolder ()
  Case "Edit"
	editFolder
  	CallFolderView()
  Case "Move"
	moveFolderForm ()
  	CallFolderView()
  Case "MoveSave"
	saveMoveFolder ()
  Case Else
	CallFolderView()
End Select
%>
<%Function CallFolderView()%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" sytle="line-height:150%">【管理新闻类别】</th>
  </tr>
  <tr>
    <td align="center" nowrap class="forumRow"><a href="NewsSort.asp?Action=Add&ParentID=0">添加一级分类</a> | <a href="NewsList.asp">查看所有新闻</a></td>
  </tr>
  <tr>
    <td nowrap class="forumRow"><%Folder(0)%></td>
  </tr>
</table>
<%
End Function
Function Folder(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_NewsSort where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("<center>暂无新闻分类</center>")
    response.end
  end if
  i=1
  response.write("<table border='0' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from Qianbo_NewsSort where ParentID="&rs("id"))(0)
    if ChildCount=0 then
	  if i=rs.recordcount then
	    FolderType="SortFileEnd"
	  else
	    FolderType="SortFile"
	  end if
	  FolderName=rs("SortName")
	  onMouseUp=""
    else
	  if i=rs.recordcount then
	 	FolderType="SortEndFolderClose"
		ListType="SortEndListline"
		onMouseUp="EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  else
		FolderType="SortFolderClose"
		ListType="SortListline"
		onMouseUp="SortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  end if
	  FolderName=rs("SortName")
    end If
    datafrom="Qianbo_NewsSort"
    response.write("<tr>")
    response.write("<td nowrap id='b"&rs("id")&"' class='"&FolderType&"'></td><td nowrap>"&FolderName&"&nbsp;")
	if rs("ViewFlag") then
      Response.Write "<a href=""Conversion.asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=down""><font color='blue'>(生效)</font></a>"
    else
      Response.Write "<a href=""Conversion.asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=up""><font color='red'>(未生效)</font></a>"
	end if
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'>操作：</font><a href='NewsSort.asp?Action=Add&ParentID="&rs("id")&"'>添加</a>")
    response.write(" | <a href='NewsSort.asp?Action=Edit&ID="&rs("id")&"'>修改</a>")
    response.write(" | <a href='NewsSort.asp?Action=Move&ID="&rs("id")&"&ParentID="&rs("Parentid")&"&SortName="&rs("SortName")&"&SortPath="&rs("SortPath")&"'>移</a>")
    response.write("→<a href='#' onclick='SortFromTo.rows[4].cells[0].innerHTML=""→ "&rs("SortName")&""";MoveForm.toID.value="&rs("ID")&";MoveForm.toParentID.value="&rs("ParentID")&";MoveForm.toSortPath.value="""&rs("SortPath")&""";'>至</a>")
	response.write(" | <a href=javascript:ConfirmDelSort('NewsSort',"&rs("id")&")>删除</a>")
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;<font color='red'>新闻：</font><a href='NewsEdit.asp?Result=Add'>添加</a>")
    response.write(" | <a href='NewsList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"'>列表</a>")
    response.write("</td></tr>")
    if ChildCount>0 then
%>
<tr id="a<%= rs("id")%>" style="display:yes">
  <td class="<%= ListType%>" nowrap></td>
  <td ><% Folder(rs("id")) %></td>
</tr>
<%
	end if
    rs.movenext
    i=i+1
	wend
	response.write("</table>")
	rs.close
	set rs=nothing
end Function

Function addFolder()
  Dim ParentID
  ParentID=request.QueryString("ParentID")
  addFolderForm ParentID
end Function

Function addFolderForm(ParentID)
  Dim ParentPath,SortTextPath,rs,sql
  if ParentID=0 then
    ParentPath="0,"
	SortTextPath=""
  else
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From Qianbo_NewsSort where ID="&ParentID
    rs.open sql,conn,1,1
	ParentPath=rs("SortPath")
  end if
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="FolderForm" method="post" action="NewsSort.asp?Action=Save&From=Add">
    <tr>
      <th height="22" sytle="line-height:150%">【添加新闻类别】</th>
    </tr>
    <tr>
      <td class="forumRow">| 根类 →
        <% if ParentID<>0 then TextPath(ParentID)%></td>
    </tr>
    <tr>
      <td class="forumRow">名称：
        <input name="SortName" type="text" id="SortName" size="28">
        生效：
        <input name="ViewFlag" type="radio" value="1" checked="checked" />
        是
        <input name="ViewFlag" type="radio" value="0" />
        否 父类ID：
        <input readonly name="ParentID" type="text" id="ParentID" size="6" value="<%=ParentID %>">
        父类数字路径：
        <input readonly name="ParentPath" type="text" id="ParentPath" size="18" value="<%=ParentPath%>">
        <input name="submitSave" type="submit" id="保存" value="保存"></td>
    </tr>
  </form>
</table>
<%
End Function
Function TextPath(ID)
  Dim rs,sql,SortTextPath
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortTextPath=rs("SortName")&"&nbsp;→&nbsp;"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(SortTextPath)
End Function
Function saveFolder
  if len(trim(request.Form("SortName")))=0 then
      response.write ("<script language='javascript'>alert('请填写类别名称！');history.back(-1);</script>")
      response.end
  end if
  Dim From,Action,rs,sql,SortTextPath
  From=request.QueryString("From")
  Set rs=server.CreateObject("adodb.recordset")
  if From="Add" then
    sql="Select * From Qianbo_NewsSort"
    rs.open sql,conn,1,3
    rs.addnew
	Action="添加新闻类别"
    rs("SortPath")=request.Form("ParentPath") & rs("ID") &","
  else
    sql="Select * From Qianbo_NewsSort where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	Action="修改新闻类别"
    rs("SortPath")=request.Form("SortPath")
  end if
  rs("SortName")=request.Form("SortName")
  rs("ViewFlag")=request.Form("ViewFlag")
  rs("ParentID")=request.Form("ParentID")
  rs.update
  response.write ("<script language='javascript'>alert('"&Action&"成功！');location.replace('NewsSort.asp');</script>")
End Function

Function editFolder()
  Dim ID
  ID=request.QueryString("ID")
  editFolderForm ID
end function

Function editFolderForm(ID)
  Dim SortName,ViewFlag,ParentID,SortPath,rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortName=rs("SortName")
  ViewFlag=rs("ViewFlag")
  ParentID=rs("ParentID")
  SortPath=rs("SortPath")
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="FolderForm" method="post" action="NewsSort.asp?Action=Save&From=Edit&ID=<%=ID%>">
    <tr>
      <th height="22" sytle="line-height:150%">【修改新闻类别】</th>
    </tr>
    <tr>
      <td class="forumRow">| 根类 →
        <% if ParentID<>0 then TextPath(ParentID)%></td>
    </tr>
    <tr>
      <td class="forumRow">名称：
        <input name="SortName" type="text" id="SortName" size="28" value="<%=SortName%>">
        发布：
        <input name="ViewFlag" type="radio" value="1" <%if ViewFlag then response.write ("checked=checked")%> />
        是
        <input name="ViewFlag" type="radio" value="0" <%if not ViewFlag then response.write ("checked=checked")%>/>
        否 父类ID：
        <input readonly name="ParentID" type="text" id="ParentID" size="6" value="<%=ParentID %>">
        父类数字路径：
        <input readonly name="SortPath" type="text" id="SortPath" size="18" value="<%=SortPath%>">
        <input name="submitSave" type="submit" id="保存" value="保存"></td>
    </tr>
  </form>
</table>
<%
End Function

Function moveFolderForm()
  Dim ID,ParentID,SortName,SortPath
  ID=request.QueryString("ID")
  ParentID=request.QueryString("ParentID")
  SortName=request.QueryString("SortName")
  SortPath=request.QueryString("SortPath")
%>
<br />
<table id="SortFromTo" class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="MoveForm" method="post" action="NewsSort.asp?Action=MoveSave">
    <tr>
      <th height="22" sytle="line-height:150%">【移动新闻类别】</th>
    </tr>
    <tr>
      <td class="forumRow">→
        <% response.write (SortName) %></td>
    </tr>
    <tr>
      <td class="forumRow">移动类ID：
        <input readonly name="ID" type="text" id="ID" size="8" value="<%=ID%>">
        移动类父ID：
        <input readonly name="ParentID" type="text" id="ParentID" size="8" value="<%=ParentID%>">
        移动类数字路径：
        <input readonly name="SortPath" type="text" id="SortPath" size="28" value="<%=SortPath%>">
        </th>
    </tr>
    <tr>
      <td align="center" class="forumRow"><strong>目标位置：通过点击"至"选择将要放置到的类别。</strong></td>
    </tr>
    <tr>
      <td class="forumRow">→ 请选择…</td>
    </tr>
    <tr>
      <td class="forumRow">目标类ID：
        <input readonly name="toID" type="text" id="toID" size="8" value="">
        目标类父ID：
        <input readonly name="toParentID" type="text" id="toParentID" size="8" value="">
        目标类数字路径：
        <input readonly name="toSortPath" type="text" id="toSortPath" size="28" value=""></td>
    </tr>
    <tr>
      <td align="center" class="forumRow"><input name="submitMove" type="submit" id="转移" value="转移">
        </th>
    </tr>
  </form>
</table>
<%
End Function

Function saveMoveFolder()
  Dim rs,sql,fromID,fromParentID,fromSortPath,toID,toParentID,toSortPath,fromParentSortPath
  fromID=request.Form("ID")
  fromParentID=request.Form("ParentID")
  fromSortPath=request.Form("SortPath")
  toID=request.Form("toID")
  toParentID=request.Form("toParentID")
  toSortPath=request.Form("toSortPath")
  if toID="" or toParentID="" or toSortPath="" then
    response.write ("<script language='javascript'>alert('请选择移动的目标位置！');history.back(-1);</script>")
    response.end
  end if
  if fromParentID=0 then
    response.write ("<script language='javascript'>alert('一级分类无法被移动！');history.back(-1);</script>")
    response.end
  end if
  if fromSortPath=toSortPath then
    response.write ("<script language='javascript'>alert('当前选择的移动类别和目标位置相同，操作无效！');history.back(-1);</script>")
    response.end
  end if
  if Instr(toSortPath,fromSortPath)>0 or fromParentID=toID then
    response.write ("<script language='javascript'>alert('不能将类别移动到本类或下属类里，操作无效！');history.back(-1);</script>")
    response.end
  end if
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_NewsSort where ID="&fromParentID
  rs.open sql,conn,0,1
  fromParentSortPath=rs("SortPath")
  conn.execute("update Qianbo_NewsSort set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")
  conn.execute("update Qianbo_NewsSort set ParentID='"&toID&"' where ID="&fromID)
  conn.execute("update Qianbo_News set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")
  response.write ("<script language='javascript'>alert('新闻类别移动成功！');location.replace('NewsSort.asp');</script>")
End Function
%>