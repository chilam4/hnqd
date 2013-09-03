<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|4,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
End If
dim Result
Result=request.QueryString("Result")
dim ID,LinkName,ViewFlag,LinkType,LinkFace,LinkUrl,Remark
ID=request.QueryString("ID")
call FriendLinkEdit()
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<script language="javascript">
<!--
function setpic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传图片，界面预览图为空，用户可以重新上传图片！");
    }
    if (arr !=null){
        editForm.LinkFace.value=arr;
    }
}
//-->
</script>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="editForm" method="post" action="FriendLinkEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>友情链接】</th>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">站点名称：</td>
    <td width="80%" class="forumRowHighlight"><input name="LinkName" type="text" id="LinkName" style="width: 180" value="<%=LinkName%>"> <input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>是否发布 <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">链接类型：</td>
    <td class="forumRowHighlight"><input name="LinkType" type="radio" value="1" <%if LinkType then response.write ("checked=checked")%> />图片 <input name="LinkType" type="radio" value="0" <%if not LinkType then response.write ("checked=checked")%> />文字</td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">前台显示文字、图片：</td>
    <td width="80%" class="forumRowHighlight"><input name="LinkFace" type="text" id="LinkFace" style="width: 280" value="<%=LinkFace%>"> <input type="button" value="上传图片" onClick="setpic();"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">链接网址：</td>
    <td class="forumRowHighlight"><input name="LinkUrl" type="text" id="LinkUrl" style="width: 280" value="<%=LinkUrl%>"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">简短说明：</td>
    <td width="80%" class="forumRowHighlight"><textarea name="Remark" rows="8" id="Remark" style="width: 500"><%=Remark%></textarea></td>
  </tr>
  <tr>
    <td align="right" class="forumRow"></td>
    <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
  </tr>
  </form>
</table>
<%
sub FriendLinkEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("LinkName")))<4 then
      response.write ("<script language='javascript'>alert('请填写网站名称并保持至少在两个汉字以上！');history.back(-1);</script>")
      response.end
    end if
    if trim(request.Form("LinkFace"))="" then
      response.write ("<script language='javascript'>alert('请填写前台显示文字或上传友情链接LOGO图片！');history.back(-1);</script>")
      response.end
    end if
    if request.Form("LinkType")=0 then
      if trim(request.Form("LinkFace"))="" then
      response.write ("<script language='javascript'>alert('请填写前台显示文字或图片地址！');history.back(-1);</script>")
      response.end
      end if
    end if
    if trim(request.Form("LinkUrl"))="" then
      response.write ("<script language='javascript'>alert('请填写友情链接网址！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then
	  sql="select * from Qianbo_FriendLink"
      rs.open sql,conn,1,3
      rs.addnew
      rs("LinkName")=trim(Request.Form("LinkName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("LinkFace")=trim(Request.Form("LinkFace"))
      rs("LinkUrl")=trim(Request.Form("LinkUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("AddTime")=now()
	end if
	if Result="Modify" then
      sql="select * from Qianbo_FriendLink where ID="&ID
      rs.open sql,conn,1,3
      rs("LinkName")=trim(Request.Form("LinkName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("LinkFace")=trim(Request.Form("LinkFace"))
      rs("LinkUrl")=trim(Request.Form("LinkUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('设置成功！');location.replace('FriendLinkList.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_FriendLink where ID="& ID
      rs.open sql,conn,1,1
	  LinkName=rs("LinkName")
	  ViewFlag=rs("ViewFlag")
	  LinkType=rs("LinkType")
	  LinkFace=rs("LinkFace")
      LinkUrl=rs("LinkUrl")
      Remark=rs("Remark")
	  rs.close
      set rs=nothing
	end if
  end if
end sub
%>