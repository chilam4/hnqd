<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|2,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end If
dim Result
Result=request.QueryString("Result")
dim ID,NavName,ViewFlag,NavUrl,HtmlNavUrl,OutFlag,Remark
ID=request.QueryString("ID")
call NavEdit()
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="editForm" method="post" action="NavigationEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>导航】</th>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">导航名称：</td>
    <td width="80%" class="forumRowHighlight"><input name="NavName" type="text" id="NavName" style="width: 200" value="<%=NavName%>" maxlength="100"> 发布：<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>> <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">动态页链接网址：</td>
    <td class="forumRowHighlight"><input name="NavUrl" type="text" id="NavUrl" style="width: 500" value="<%=NavUrl%>"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">静态页链接网址：</td>
    <td class="forumRowHighlight"><input name="HtmlNavUrl" type="text" id="HtmlNavUrl" style="width: 500" value="<%=HtmlNavUrl%>"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">链接状态：</td>
    <td width="80%" class="forumRowHighlight"><input name="OutFlag" type="checkbox" value="1" <%if OutFlag then response.write ("checked")%>>是否外部链接</td>
  </tr>
  <tr>
    <td align="right" class="forumRow">备注：</td>
    <td class="forumRowHighlight"><textarea name="Remark" rows="8" id="Remark" style="width: 500"><%=Remark%></textarea></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow"></td>
    <td width="80%" class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
  </tr>
  </form>
</table>
<%
sub NavEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NavName")))<2 then
		response.write ("<script language='javascript'>alert('请填写导航名称并保持至少在一个汉字以上！');history.back(-1);</script>")
		response.end
    end If
	If trim(Request.Form("NavName")) = ""  Or trim(Request.Form("NavUrl")) = ""  Or trim(Request.Form("HtmlNavUrl")) = "" Then
		response.write ("<script language='javascript'>alert('请填写导航名称和链接网址！');history.back(-1);</script>")
		response.end
	End If
    if Result="Add" Then
	  sql="select * from Qianbo_Navigation"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  rs("HtmlNavUrl")=trim(Request.Form("HtmlNavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("OutFlag")=1 then
        rs("OutFlag")=Request.Form("OutFlag")
	  else
        rs("OutFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("Sequence")=99
	  rs("AddTime")=now()
	end if
	if Result="Modify" Then
      sql="select * from Qianbo_Navigation where ID="&ID
      rs.open sql,conn,1,3
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  rs("HtmlNavUrl")=trim(Request.Form("HtmlNavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("OutFlag")=1 then
        rs("OutFlag")=Request.Form("OutFlag")
	  else
        rs("OutFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('设置成功！');location.replace('NavigationList.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Navigation where ID="& ID
      rs.open sql,conn,1,1
	  NavName=rs("NavName")
	  ViewFlag=rs("ViewFlag")
      Remark=rs("Remark")
	  OutFlag=rs("OutFlag")
      NavUrl=rs("NavUrl")
	  HtmlNavUrl=rs("HtmlNavUrl")
	  rs.close
      set rs=nothing
	end if
  end if
end sub
%>