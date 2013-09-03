<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Admin_htmlconfig.asp"-->
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
        editForm.SmallPic.value=arr;
    }
}
function setbpic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传图片，界面预览图为空，用户可以重新上传图片！");
    }
    if (arr !=null){
        editForm.BigPic.value=arr;
    }
}
function SetOtherPic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传图片，界面预览图为空，用户可以重新上传图片！");
    }
    if (arr !=null){
        editForm.OtherPic.value+='*'+arr;
    }
}
//-->
</script>
<%
if Instr(session("AdminPurview"),"|13,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
dim Result
Result=request.QueryString("Result")
dim ID,ProductName,ViewFlag,SortName,SortID,SortPath
dim ProductNo,ProductModel,N_Price,P_Price,Stock,Unit,Maker,CommendFlag,NewFlag,GroupID,GroupIdName,Exclusive,SeoKeywords,SeoDescription
dim SmallPic,BigPic,OtherPic,Content
ID=request.QueryString("ID")
call ProductEdit()
call SiteInfo
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="ProductEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>产品】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">产品标题：</td>
      <td width="80%" class="forumRowHighlight"><input name="ProductName" type="text" id="ProductName" style="width: 280" value="<%=ProductName%>" maxlength="250">
        是否生效：
        <input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>
        是否推荐：
        <input name="CommendFlag" type="checkbox" style="height: 13px;width: 13px;" value="1" <%if CommendFlag then response.write ("checked")%>>
        是否新品：
        <input name="NewFlag" type="checkbox" value="1" style="height: 13px;width: 13px;" <%if NewFlag then response.write ("checked")%>>
        <font color="red">*</font> <input type="button" name="btn" value="复制标题" title="复制标题到：MetaDescription、MetaKeywords" onclick="CopyWebTitle(document.editForm.ProductName.value);"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">MetaKeywords：</td>
      <td class="forumRowHighlight"><input name="SeoKeywords" type="text" id="SeoKeywords" style="width: 500" value="<%=SeoKeywords%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">MetaDescription：</td>
      <td class="forumRowHighlight"><input name="SeoDescription" type="text" id="SeoDescription" style="width: 500" value="<%=SeoDescription%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品类别：</td>
      <td class="forumRowHighlight"><input name="SortID" type="text" id="SortID" style="width: 18; background-color:#fffff0" value="<%=SortID%>" readonly>
        <input name="SortPath" type="text" id="SortPath" style="width: 70; background-color:#fffff0" value="<%=SortPath%>" readonly>
        <input name="SortName" type="text" id="SortName" value="<%=SortName%>" style="width: 180; background-color:#fffff0" readonly>
        <a href="javaScript:OpenScript('SelectSort.asp?Result=Products',500,500,'')">选择类别</a> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品编号：</td>
      <td class="forumRowHighlight"><input name="ProductNo" type="text" style="width: 180;" value="<%=ProductNo%>" maxlength="180">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品型号：</td>
      <td class="forumRowHighlight"><input name="ProductModel" type="text" style="width: 180;" value="<%=ProductModel%>" maxlength="180">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">市场价格：</td>
      <td class="forumRowHighlight"><input name="N_Price" type="text" style="width: 80;" value="<%=N_Price%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">批发价格：</td>
      <td class="forumRowHighlight"><input name="P_Price" type="text" style="width: 80;" value="<%=P_Price%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">库存数量：</td>
      <td class="forumRowHighlight"><input name="Stock" type="text" style="width: 80;" value="<%=Stock%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">计价单位：</td>
      <td class="forumRowHighlight"><input name="Unit" type="text" style="width: 80;" value="<%=Unit%>" maxlength="80">
        <a href="javascript:" onClick="document.editForm.Unit.value='件'">件</a> <a href="javascript:" onClick="document.editForm.Unit.value='套'">套</a> <a href="javascript:" onClick="document.editForm.Unit.value='个'">个</a> <a href="javascript:" onClick="document.editForm.Unit.value='只'">只</a> <a href="javascript:" onClick="document.editForm.Unit.value='台'">台</a> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">出品公司：</td>
      <td class="forumRowHighlight"><input name="Maker" type="text" style="width: 250;" value="<%=Maker%>" maxlength="250">
        <a href="javascript:" onClick="document.editForm.Maker.value='<%=SiteTitle%>'"><%=SiteTitle%></a> <font color="red">*</font></td>
    </tr>
<%
if Result="Modify" then
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Products where ID="& ID
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write ("<center>数据库记录读取错误！</center>")
	response.end
end If
if rs("attribute1")<>"" and rs("attribute1_value")<>"" then
	attribute1_1=Split(rs("attribute1"),"§§§")
	attribute1_value_1=Split(rs("attribute1_value"),"§§§")
	Num_1=ubound(attribute1_1)+1
Else
	Num_1=0
End If
rs.close
set rs=Nothing
Else
	Num_1=0
End If
%>
    <tr>
      <td align="right" class="forumRow">自定义产品属性：</td>
      <td class="forumRowHighlight"><input name="Num_1" type="text" id="Num_1" value="<%=Num_1%>" size="5" />个 <input name="button2" type="button" id="button2" value="设置" onClick="num_1()" /> <input type="button" name="button7" id="button7" value="增加一个" onClick="num_1_1()" />
        <br />
        <span id="num_1_str">
        <%For i=0 to (Num_1-1)%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="28">属性名称：
              <input name="attribute<%=i+1%>" type="text" id="attribute<%=i+1%>" value="<%=attribute1_1(i)%>" size="18" />
              属性值：
              <input name="attribute<%=i+1%>_value" type="text" id="attribute<%=i+1%>_value" value="<%=attribute1_value_1(i)%>" size="50" /></td>
          </tr>
        </table>
        <%Next%>
        </span> </td>
    </tr>
    </tr>
    
    <tr>
      <td align="right" class="forumRow">阅读权限：</td>
      <td class="forumRowHighlight"><select name="GroupID">
          <% call SelectGroup() %>
        </select>
        <input name="Exclusive" type="radio" value="&gt;=" <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>>
        隶属
        <input type="radio" <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
        专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品小图：</td>
      <td class="forumRowHighlight"><input name="SmallPic" type="text" style="width: 280;" value="<%=SmallPic%>" maxlength="250">
        <input type="button" value="上传图片" onClick="setpic();">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品大图：</td>
      <td class="forumRowHighlight"><input name="BigPic" type="text" style="width: 280;" value="<%=BigPic%>" maxlength="250">
        <input type="button" value="上传图片" onClick="setbpic();">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">更多图片：</td>
      <td class="forumRowHighlight"><textarea rows="5" cols="80" name="OtherPic"><%=OtherPic%></textarea>
        <input type="button" value="上传图片" onClick="SetOtherPic();"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品内容：</td>
      <td class="forumRowHighlight"><textarea name="Content" id="Content" style="display:none"><%=Server.HTMLEncode((Content))%></textarea>
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=Content&style=coolblue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存">
        <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
sub ProductEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ProductName")))<1 then
      response.write ("<script language='javascript'>alert('请填写产品名称！');history.back(-1);</script>")
      response.end
    end If
	if Request.Form("SortID")="" and Request.Form("SortPath")="" then
		response.write ("<script language='javascript'>alert('请选择所属分类！');history.back(-1);</script>")
		response.End
	end If
	if ltrim(request.Form("ProductModel")) = "" then
		response.write ("<script language='javascript'>alert('请填写产品型号！');history.back(-1);</script>")
		response.end
	end If
	if (not IsNumeric(trim(request.Form("N_Price")))) or (not IsNumeric(trim(request.Form("P_Price"))))then
		response.write ("<script language='javascript'>alert('请正确填写市场价格、批发价格！');history.back(-1);</script>")
		response.end
	elseif trim(request.Form("N_Price"))<0 or trim(request.Form("P_Price"))<0then
		response.write ("<script language='javascript'>alert('请正确填写市场价格、批发价格！');history.back(-1);</script>")
		response.end
	end If
	if (not IsNumeric(trim(request.Form("Stock"))))  then
		response.write ("<script language='javascript'>alert('请填写/选择库存数量！');history.back(-1);</script>")
		response.end
	end if
	if len(trim(Request.Form("Unit")))=0 then
		response.write ("<script language='javascript'>alert('请填写/选择产品单位！');history.back(-1);</script>")
		response.end
	end If
    if ltrim(request.Form("Maker")) = "" then
		response.write ("<script language='javascript'>alert('请填写出品公司！');history.back(-1);</script>")
		response.end
    end If
	if ltrim(request.Form("SmallPic")) = "" then
		response.write ("<script language='javascript'>alert('请上传产品小图！');history.back(-1);</script>")
		response.end
	end If
	if ltrim(request.Form("BigPic")) = "" then
		response.write ("<script language='javascript'>alert('请上传产品大图！');history.back(-1);</script>")
		response.end
	end If
	if ltrim(request.Form("Content")) = "" then
		response.write ("<script language='javascript'>alert('请填写产品详细介绍！');history.back(-1);</script>")
		response.end
	end If
    if Result="Add" Then
	  set rsRepeat = conn.execute("select ProductNo from Qianbo_Products where ProductNo='" & trim(Request.Form("ProductNo")) & "'")
	  if not (rsRepeat.bof and rsRepeat.eof) then
		response.write "<script language='javascript'>alert('" & trim(Request.Form("ProductNo")) & "产品编号已存在！');history.back(-1);</script>"
		response.End
	  End If
	  rsRepeat.close
	  set rsRepeat=Nothing
	  sql="select * from Qianbo_Products"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ProductName")=trim(Request.Form("ProductName"))
      if Request.Form("ViewFlag")=1 then
		rs("ViewFlag")=Request.Form("ViewFlag")
      else
		rs("ViewFlag")=0
      end if
      rs("SortID")=Request.Form("SortID")
      rs("SortPath")=Request.Form("SortPath")
      rs("ProductNo")=trim(Request.Form("ProductNo"))
      rs("ProductModel")=trim(Request.Form("ProductModel"))
      rs("N_Price")=Round(trim(Request.Form("N_Price")),2)
      rs("P_Price")=Round(trim(Request.Form("P_Price")),2)
      rs("Stock")=Round(trim(Request.Form("Stock")),2)
      rs("Unit")=trim(Request.Form("Unit"))
      rs("Maker")=trim(Request.Form("Maker"))
      if Request.Form("CommendFlag")=1 then
		rs("CommendFlag")=Request.Form("CommendFlag")
      else
		rs("CommendFlag")=0
      end if
      if Request.Form("NewFlag")=1 then
		rs("NewFlag")=Request.Form("NewFlag")
      else
		rs("NewFlag")=0
      end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
      rs("GroupID")=GroupIdName(0)
      rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
      rs("BigPic")=trim(Request.Form("BigPic"))
	  rs("OtherPic")=trim(Request.Form("OtherPic"))
      rs("Content")=rtrim(Request.Form("Content"))
      rs("AddTime")=now()
      rs("UpdateTime")=now()
      Num_1=CheckStr(Request.Form("Num_1"),1)
      if Num_1="" then Num_1=0
      if Num_1>0 then
		For i=1 to Num_1
			If CheckStr(Request.Form("attribute"&i),0)<>"" and  CheckStr(Request.Form("attribute"&i&"_value"),0)<>"" Then
				If attribute1="" then
					attribute1=CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=CheckStr(Request.Form("attribute"&i&"_value"),0)
				Else
					attribute1=attribute1&"§§§"&CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=attribute1_value&"§§§"&CheckStr(Request.Form("attribute"&i&"_value"),0)
				End if
			End If
		Next
      end if
	  rs("attribute1")=attribute1
	  rs("attribute1_value")=attribute1_value
	  rs("SeoKeywords")=trim(Request.Form("SeoKeywords"))
	  rs("SeoDescription")=trim(Request.Form("SeoDescription"))
	  rs.update
	  rs.close
	  set rs=Nothing
	  set rs=server.createobject("adodb.recordset")
	  sql="select top 1 ID from Qianbo_Products order by ID desc"
	  rs.open sql,conn,1,1
	  ID=rs("ID")
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&ProName&""&Separated&""&ID&"."&HTMLName&"","ProductView.asp","ID=",ID,"","")
	  End If
	  End If
	  if Result="Modify" then
      sql="select * from Qianbo_Products where ID="&ID
      rs.open sql,conn,1,3
      rs("ProductName")=trim(Request.Form("ProductName"))
	  if Request.Form("ViewFlag")=1 then
		rs("ViewFlag")=Request.Form("ViewFlag")
	  else
		rs("ViewFlag")=0
	  end if
	  rs("SortID")=Request.Form("SortID")
	  rs("SortPath")=Request.Form("SortPath")
	  rs("ProductNo")=trim(Request.Form("ProductNo"))
	  rs("ProductModel")=trim(Request.Form("ProductModel"))
	  rs("N_Price")=Round(trim(Request.Form("N_Price")),2)
	  rs("P_Price")=Round(trim(Request.Form("P_Price")),2)
	  rs("Stock")=Round(trim(Request.Form("Stock")),2)
	  rs("Unit")=trim(Request.Form("Unit"))
	  rs("Maker")=trim(Request.Form("Maker"))
	  if Request.Form("CommendFlag")=1 then
		rs("CommendFlag")=Request.Form("CommendFlag")
	  else
		rs("CommendFlag")=0
	  end if
	  if Request.Form("NewFlag")=1 then
		rs("NewFlag")=Request.Form("NewFlag")
	  else
		rs("NewFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("BigPic")=trim(Request.Form("BigPic"))
	  rs("OtherPic")=trim(Request.Form("OtherPic"))
	  rs("Content")=rtrim(Request.Form("Content"))
	  rs("UpdateTime")=now()
	  Num_1=CheckStr(Request.Form("Num_1"),1)
	  if Num_1="" then Num_1=0
	  if Num_1>0 then
		For i=1 to Num_1
			If CheckStr(Request.Form("attribute"&i),0)<>"" and  CheckStr(Request.Form("attribute"&i&"_value"),0)<>"" Then
				If attribute1="" then
					attribute1=CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=CheckStr(Request.Form("attribute"&i&"_value"),0)
				Else
					attribute1=attribute1&"§§§"&CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=attribute1_value&"§§§"&CheckStr(Request.Form("attribute"&i&"_value"),0)
				End if
			End If
		Next
	  end if
	  rs("attribute1")=attribute1
	  rs("attribute1_value")=attribute1_value
	  rs("SeoKeywords")=trim(Request.Form("SeoKeywords"))
	  rs("SeoDescription")=trim(Request.Form("SeoDescription"))
	  rs.update
	  rs.close
	  set rs=Nothing
	  if ISHTML = 1 then
	  call htmll("","",""&ProName&""&Separated&""&ID&"."&HTMLName&"","ProductView.asp","ID=",ID,"","")
	  End If
	  End If
	  if ISHTML = 1 then
	  response.write "<script language='javascript'>alert('设置成功，相关静态页面已更新！');location.replace('ProductList.asp');</script>"
	  Else
	  response.write "<script language='javascript'>alert('设置成功！');location.replace('ProductList.asp');</script>"
	  End If
  else
  	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Products where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
      response.write ("<center>数据库记录读取错误！</center>")
      response.end
      end if
      ProductName=rs("ProductName")
      ViewFlag=rs("ViewFlag")
      SortName=SortText(rs("SortID"))
      SortID=rs("SortID")
      SortPath=rs("SortPath")
      ProductNo=rs("ProductNo")
      ProductModel=rs("ProductModel")
      N_Price=rs("N_Price")
      P_Price=rs("P_Price")
      Stock=rs("Stock")
      Unit=rs("Unit")
      Maker=rs("Maker")
      CommendFlag=rs("CommendFlag")
      NewFlag=rs("NewFlag")
      GroupID=rs("GroupID")
      Exclusive=rs("Exclusive")
	  SmallPic=rs("SmallPic")
      BigPic=rs("BigPic")
	  OtherPic=rs("OtherPic")
      Content=rs("Content")
	  SeoKeywords=rs("SeoKeywords")
	  SeoDescription=rs("SeoDescription")
      rs.close
      set rs=nothing
	  else
      randomize timer
      ProductNo=Hour(now)&Minute(now)&Second(now)&"-"&int(900*rnd)+100
      Stock=10000
    end if
  end if
end sub

sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from Qianbo_MemGroup"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
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
  sql="Select * From Qianbo_ProductSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>