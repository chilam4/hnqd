<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
if Instr(session("AdminPurview"),"|29,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
dim Result
Result=request.QueryString("Result")
dim ID,AdminName,Working,Password,vPassword,UserName,Purview,Explain,AddTime
ID=request.QueryString("ID")
if ID="" then ID=0
call AdminEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="AdminEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>管理员】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">登录名称：</td>
      <td width="80%" class="forumRowHighlight"><input name="AdminName" type="text" id="AdminName" style="width: 180" value="<%=AdminName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>
        <font color="red">*</font>3-10个字符</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">生效：</td>
      <td class="forumRowHighlight"><input name="Working" type="checkbox" value="1" <%if Working then response.write ("checked")%>></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">管理员密码：</td>
      <td class="forumRowHighlight"><input name="Password" type="password" id="Password" maxlength="20" style="width: 180">
        <font color="red">*</font>6-16个字符</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">确认密码：</td>
      <td class="forumRowHighlight"><input name="vPassword" type="password" id="vPassword" maxlength="20" style="width: 180">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">管理员名称：</td>
      <td class="forumRowHighlight"><input name="UserName" type="text" id="UserName" style="width: 120;" value="<%=UserName%>"></td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow">操作权限：</td>
      <td class="forumRowHighlight">
        <input name="Purview1" type="checkbox" value="|1,"<%if Instr(Purview,"|1,")>0 then response.write ("checked")%>>网站参数设置
        <input name="Purview2" type="checkbox" value="|2,"<%if Instr(Purview,"|2,")>0 then response.write ("checked")%>>导航栏添加
        <input name="Purview3" type="checkbox" value="|3,"<%if Instr(Purview,"|3,")>0 then response.write ("checked")%>>导航栏管理
        <input name="Purview4" type="checkbox" value="|4,"<%if Instr(Purview,"|4,")>0 then response.write ("checked")%>>友情链接添加
        <input name="Purview5" type="checkbox" value="|5,"<%if Instr(Purview,"|5,")>0 then response.write ("checked")%>>友情链接管理
        <input name="Purview6" type="checkbox" value="|6,"<%if Instr(Purview,"|6,")>0 then response.write ("checked")%>>新闻类别管理</td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview7" type="checkbox" value="|7,"<%if Instr(Purview,"|7,")>0 then response.write ("checked")%>>新闻列表管理
        <input name="Purview8" type="checkbox" value="|8,"<%if Instr(Purview,"|8,")>0 then response.write ("checked")%>>添加新闻
        <input name="Purview9" type="checkbox" value="|9,"<%if Instr(Purview,"|9,")>0 then response.write ("checked")%>>企业信息列表
        <input name="Purview10" type="checkbox" value="|10,"<%if Instr(Purview,"|10,")>0 then response.write ("checked")%>>添加企业信息
        <input name="Purview11" type="checkbox" value="|11,"<%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>产品类别管理
        <input name="Purview12" type="checkbox" value="|12,"<%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>产品列表管理</td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview13" type="checkbox" value="|13,"<%if Instr(Purview,"|13,")>0 then response.write ("checked")%>>添加产品信息
        <input name="Purview14" type="checkbox" value="|14,"<%if Instr(Purview,"|14,")>0 then response.write ("checked")%>>下载类别管理
        <input name="Purview15" type="checkbox" value="|15,"<%if Instr(Purview,"|15,")>0 then response.write ("checked")%>>下载列表管理
        <input name="Purview16" type="checkbox" value="|16,"<%if Instr(Purview,"|16,")>0 then response.write ("checked")%>>添加下载信息
        <input name="Purview17" type="checkbox" value="|17,"<%if Instr(Purview,"|17,")>0 then response.write ("checked")%>>招聘列表管理
        <input name="Purview18" type="checkbox" value="|18,"<%if Instr(Purview,"|18,")>0 then response.write ("checked")%>>添加招聘信息</td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview19" type="checkbox" value="|19,"<%if Instr(Purview,"|19,")>0 then response.write ("checked")%>>信息类别管理
        <input name="Purview20" type="checkbox" value="|20,"<%if Instr(Purview,"|20,")>0 then response.write ("checked")%>>信息列表管理
        <input name="Purview21" type="checkbox" value="|21,"<%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>添加信息
        <input name="Purview22" type="checkbox" value="|22,"<%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>留言信息查看
		<input name="Purview23" type="checkbox" value="|23,"<%if Instr(Purview,"|23,")>0 then response.write ("checked")%>>留言信息管理
        <input name="Purview24" type="checkbox" value="|24,"<%if Instr(Purview,"|24,")>0 then response.write ("checked")%>>订单信息查看</td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview25" type="checkbox" value="|25,"<%if Instr(Purview,"|25,")>0 then response.write ("checked")%>>订单信息管理
        <input name="Purview26" type="checkbox" value="|26,"<%if Instr(Purview,"|26,")>0 then response.write ("checked")%>>人才信息查看
        <input name="Purview27" type="checkbox" value="|27,"<%if Instr(Purview,"|27,")>0 then response.write ("checked")%>>人才信息管理
		<input name="Purview28" type="checkbox" value="|28,"<%if Instr(Purview,"|28,")>0 then response.write ("checked")%>>网站管理员查看
        <input name="Purview29" type="checkbox" value="|29,"<%if Instr(Purview,"|29,")>0 then response.write ("checked")%>>网站管理员管理
		<input name="Purview30" type="checkbox" value="|30,"<%if Instr(Purview,"|30,")>0 then response.write ("checked")%>>会员资料查看</td>
    </tr>
    <tr <%if ID=1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview31" type="checkbox" value="|31,"<%if Instr(Purview,"|31,")>0 then response.write ("checked")%>>会员资料管理
		<input name="Purview32" type="checkbox" value="|32,"<%if Instr(Purview,"|32,")>0 then response.write ("checked")%>>会员组别管理
		<input name="Purview33" type="checkbox" value="|33,"<%if Instr(Purview,"|33,")>0 then response.write ("checked")%>>后台登录日志管理
		<input name="Purview34" type="checkbox" value="|34,"<%if Instr(Purview,"|34,")>0 then response.write ("checked")%>>生成静态页面管理
		<input name="Purview35" type="checkbox" value="|35,"<%if Instr(Purview,"|35,")>0 then response.write ("checked")%>>编辑器管理</td>
    </tr>
	<tr <%if ID=1 then response.write ("style=display:none")%>>
      <td class="forumRow"></td>
      <td class="forumRowHighlight">
        <input name="Purview36" type="checkbox" value="|36,"<%if Instr(Purview,"|36,")>0 then response.write ("checked")%>>站内链接管理
        <input name="Purview37" type="checkbox" value="|37,"<%if Instr(Purview,"|37,")>0 then response.write ("checked")%>>站内链接添加
		<input name="Purview38" type="checkbox" value="|38,"<%if Instr(Purview,"|38,")>0 then response.write ("checked")%>>数据表查看
		<input name="Purview39" type="checkbox" value="|39,"<%if Instr(Purview,"|39,")>0 then response.write ("checked")%>>数据表管理
		<input name="Purview39" type="checkbox" value="|39,"<%if Instr(Purview,"|40,")>0 then response.write ("checked")%>>客户即时咨询管理
        <input name="Purview39" type="checkbox" value="|40,"<%if Instr(Purview,"|39,")>0 then response.write ("checked")%>>谷歌SiteMap</td>
    </tr>
    <tr <%if ID<>1 then response.write ("style=display:none")%>>
      <td align="right" class="forumRow">操作权限：</td>
      <td class="forumRowHighlight">内置超级管理员帐号</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">备注：</td>
      <td class="forumRowHighlight"><textarea name="Explain" rows="8" id="Explain" style="width: 500" ><%=Explain%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存">
        <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
sub AdminEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then
      set rsCheckAdd = conn.execute("select AdminName from Qianbo_Admin where AdminName='" & trim(Request.Form("AdminName")) & "'")
      if not (rsCheckAdd.bof and rsCheckAdd.eof) then
        response.write "<script language='javascript'>alert('" & trim(Request.Form("AdminName")) & "管理员名称已存在！');history.back(-1);</script>"
        response.end
      end if
	  sql="select * from Qianbo_Admin"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("AdminName")))<3 or len(trim(Request.Form("AdminName")))>10  then
        response.write "<script language='javascript'>alert('请填写管理员名称(字符数在3-10位之间)！');history.back(-1);</script>"
        response.end
      end if
      if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
        response.write "<script language='javascript'>alert('请填写管理员密码(字符数在6-16位之间)！');history.back(-1);</script>"
        response.end
      end if
	  if Request.Form("Password")<>Request.Form("vPassword") then
        response.write "<script language='javascript'>alert('两次输入的密码不同！');history.back(-1);</script>"
        response.end
	  end if
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  rs("Password")=Md5(Request.Form("Password"))
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview1") & Request.Form("Purview2") &_
	                     Request.Form("Purview3") & Request.Form("Purview4") & Request.Form("Purview5") &_
	                     Request.Form("Purview6") & Request.Form("Purview7") & Request.Form("Purview8") &_
	                     Request.Form("Purview9") & Request.Form("Purview10") & Request.Form("Purview11") &_
	                     Request.Form("Purview12") & Request.Form("Purview13") &_
	                     Request.Form("Purview14") & Request.Form("Purview15") & Request.Form("Purview16") &_
	                     Request.Form("Purview17") & Request.Form("Purview18") &_
	                     Request.Form("Purview19") & Request.Form("Purview20") & Request.Form("Purview21") &_
	                     Request.Form("Purview22") & Request.Form("Purview23") & Request.Form("Purview24") &_
	                     Request.Form("Purview25") &_
						 Request.Form("Purview26") & Request.Form("Purview27") & Request.Form("Purview28") &_
						 Request.Form("Purview29") & Request.Form("Purview30") & Request.Form("Purview31") &_
						 Request.Form("Purview32") & Request.Form("Purview33") & Request.Form("Purview34") &_
	                     Request.Form("Purview35") & Request.Form("Purview36") & Request.Form("Purview37") &_
						 Request.Form("Purview38") & Request.Form("Purview39")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if
	if Result="Modify" then
      sql="select * from Qianbo_Admin where ID="&ID
      rs.open sql,conn,1,3
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
      if trim(Request.Form("Password"))<>"" then
	    if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
          response.write "<script language='javascript'>alert('请填写管理员密码(字符数在6-16位之间)！');history.back(-1);</script>"
          response.end
        end if
	    if Request.Form("Password")<>Request.Form("vPassword") then
          response.write "<script language='javascript'>alert('两次输入的密码不同！');history.back(-1);</script>"
          response.end
	    end if
	    rs("Password")=Md5(Request.Form("Password"))
	  end if
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview1") & Request.Form("Purview2") &_
	                     Request.Form("Purview3") & Request.Form("Purview4") & Request.Form("Purview5") &_
	                     Request.Form("Purview6") & Request.Form("Purview7") & Request.Form("Purview8") &_
	                     Request.Form("Purview9") & Request.Form("Purview10") & Request.Form("Purview11") &_
	                     Request.Form("Purview12") & Request.Form("Purview13") &_
	                     Request.Form("Purview14") & Request.Form("Purview15") & Request.Form("Purview16") &_
	                     Request.Form("Purview17") & Request.Form("Purview18") &_
	                     Request.Form("Purview19") & Request.Form("Purview20") & Request.Form("Purview21") &_
	                     Request.Form("Purview22") & Request.Form("Purview23") & Request.Form("Purview24") &_
	                     Request.Form("Purview25") &_
						 Request.Form("Purview26") & Request.Form("Purview27") & Request.Form("Purview28") &_
						 Request.Form("Purview29") & Request.Form("Purview30") & Request.Form("Purview31") &_
						 Request.Form("Purview32") & Request.Form("Purview33") & Request.Form("Purview34") &_
	                     Request.Form("Purview35") & Request.Form("Purview36") & Request.Form("Purview37") &_
						 Request.Form("Purview38") & Request.Form("Purview39")
	  rs("Explain")=trim(Request.Form("Explain"))
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('设置成功！');location.replace('AdminList.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Admin where ID="& ID
      rs.open sql,conn,1,1
	  AdminName=rs("AdminName")
	  Working=rs("Working")
	  UserName=rs("UserName")
	  Purview=rs("AdminPurview")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing
	end if
  end if
end sub
%>