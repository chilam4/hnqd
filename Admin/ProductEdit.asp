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
        alert("ϵͳ��ʾ����ǰû���ϴ�ͼƬ������Ԥ��ͼΪ�գ��û����������ϴ�ͼƬ��");
    }
    if (arr !=null){
        editForm.SmallPic.value=arr;
    }
}
function setbpic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("ϵͳ��ʾ����ǰû���ϴ�ͼƬ������Ԥ��ͼΪ�գ��û����������ϴ�ͼƬ��");
    }
    if (arr !=null){
        editForm.BigPic.value=arr;
    }
}
function SetOtherPic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("ϵͳ��ʾ����ǰû���ϴ�ͼƬ������Ԥ��ͼΪ�գ��û����������ϴ�ͼƬ��");
    }
    if (arr !=null){
        editForm.OtherPic.value+='*'+arr;
    }
}
//-->
</script>
<%
if Instr(session("AdminPurview"),"|13,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
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
      <th height="22" colspan="2" sytle="line-height:150%">��<%If Result = "Add" then%>���<%ElseIf Result = "Modify" then%>�޸�<%End If%>��Ʒ��</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">��Ʒ���⣺</td>
      <td width="80%" class="forumRowHighlight"><input name="ProductName" type="text" id="ProductName" style="width: 280" value="<%=ProductName%>" maxlength="250">
        �Ƿ���Ч��
        <input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>
        �Ƿ��Ƽ���
        <input name="CommendFlag" type="checkbox" style="height: 13px;width: 13px;" value="1" <%if CommendFlag then response.write ("checked")%>>
        �Ƿ���Ʒ��
        <input name="NewFlag" type="checkbox" value="1" style="height: 13px;width: 13px;" <%if NewFlag then response.write ("checked")%>>
        <font color="red">*</font> <input type="button" name="btn" value="���Ʊ���" title="���Ʊ��⵽��MetaDescription��MetaKeywords" onclick="CopyWebTitle(document.editForm.ProductName.value);"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">MetaKeywords��</td>
      <td class="forumRowHighlight"><input name="SeoKeywords" type="text" id="SeoKeywords" style="width: 500" value="<%=SeoKeywords%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">MetaDescription��</td>
      <td class="forumRowHighlight"><input name="SeoDescription" type="text" id="SeoDescription" style="width: 500" value="<%=SeoDescription%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ���</td>
      <td class="forumRowHighlight"><input name="SortID" type="text" id="SortID" style="width: 18; background-color:#fffff0" value="<%=SortID%>" readonly>
        <input name="SortPath" type="text" id="SortPath" style="width: 70; background-color:#fffff0" value="<%=SortPath%>" readonly>
        <input name="SortName" type="text" id="SortName" value="<%=SortName%>" style="width: 180; background-color:#fffff0" readonly>
        <a href="javaScript:OpenScript('SelectSort.asp?Result=Products',500,500,'')">ѡ�����</a> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ��ţ�</td>
      <td class="forumRowHighlight"><input name="ProductNo" type="text" style="width: 180;" value="<%=ProductNo%>" maxlength="180">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ�ͺţ�</td>
      <td class="forumRowHighlight"><input name="ProductModel" type="text" style="width: 180;" value="<%=ProductModel%>" maxlength="180">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�г��۸�</td>
      <td class="forumRowHighlight"><input name="N_Price" type="text" style="width: 80;" value="<%=N_Price%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�����۸�</td>
      <td class="forumRowHighlight"><input name="P_Price" type="text" style="width: 80;" value="<%=P_Price%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���������</td>
      <td class="forumRowHighlight"><input name="Stock" type="text" style="width: 80;" value="<%=Stock%>" maxlength="80">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�Ƽ۵�λ��</td>
      <td class="forumRowHighlight"><input name="Unit" type="text" style="width: 80;" value="<%=Unit%>" maxlength="80">
        <a href="javascript:" onClick="document.editForm.Unit.value='��'">��</a> <a href="javascript:" onClick="document.editForm.Unit.value='��'">��</a> <a href="javascript:" onClick="document.editForm.Unit.value='��'">��</a> <a href="javascript:" onClick="document.editForm.Unit.value='ֻ'">ֻ</a> <a href="javascript:" onClick="document.editForm.Unit.value='̨'">̨</a> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ��˾��</td>
      <td class="forumRowHighlight"><input name="Maker" type="text" style="width: 250;" value="<%=Maker%>" maxlength="250">
        <a href="javascript:" onClick="document.editForm.Maker.value='<%=SiteTitle%>'"><%=SiteTitle%></a> <font color="red">*</font></td>
    </tr>
<%
if Result="Modify" then
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Products where ID="& ID
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write ("<center>���ݿ��¼��ȡ����</center>")
	response.end
end If
if rs("attribute1")<>"" and rs("attribute1_value")<>"" then
	attribute1_1=Split(rs("attribute1"),"����")
	attribute1_value_1=Split(rs("attribute1_value"),"����")
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
      <td align="right" class="forumRow">�Զ����Ʒ���ԣ�</td>
      <td class="forumRowHighlight"><input name="Num_1" type="text" id="Num_1" value="<%=Num_1%>" size="5" />�� <input name="button2" type="button" id="button2" value="����" onClick="num_1()" /> <input type="button" name="button7" id="button7" value="����һ��" onClick="num_1_1()" />
        <br />
        <span id="num_1_str">
        <%For i=0 to (Num_1-1)%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="28">�������ƣ�
              <input name="attribute<%=i+1%>" type="text" id="attribute<%=i+1%>" value="<%=attribute1_1(i)%>" size="18" />
              ����ֵ��
              <input name="attribute<%=i+1%>_value" type="text" id="attribute<%=i+1%>_value" value="<%=attribute1_value_1(i)%>" size="50" /></td>
          </tr>
        </table>
        <%Next%>
        </span> </td>
    </tr>
    </tr>
    
    <tr>
      <td align="right" class="forumRow">�Ķ�Ȩ�ޣ�</td>
      <td class="forumRowHighlight"><select name="GroupID">
          <% call SelectGroup() %>
        </select>
        <input name="Exclusive" type="radio" value="&gt;=" <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>>
        ����
        <input type="radio" <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
        ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ƷСͼ��</td>
      <td class="forumRowHighlight"><input name="SmallPic" type="text" style="width: 280;" value="<%=SmallPic%>" maxlength="250">
        <input type="button" value="�ϴ�ͼƬ" onClick="setpic();">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ��ͼ��</td>
      <td class="forumRowHighlight"><input name="BigPic" type="text" style="width: 280;" value="<%=BigPic%>" maxlength="250">
        <input type="button" value="�ϴ�ͼƬ" onClick="setbpic();">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">����ͼƬ��</td>
      <td class="forumRowHighlight"><textarea rows="5" cols="80" name="OtherPic"><%=OtherPic%></textarea>
        <input type="button" value="�ϴ�ͼƬ" onClick="SetOtherPic();"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ���ݣ�</td>
      <td class="forumRowHighlight"><textarea name="Content" id="Content" style="display:none"><%=Server.HTMLEncode((Content))%></textarea>
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=Content&style=coolblue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="����">
        <input type="button" value="������һҳ" onclick="history.back(-1)"></td>
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
      response.write ("<script language='javascript'>alert('����д��Ʒ���ƣ�');history.back(-1);</script>")
      response.end
    end If
	if Request.Form("SortID")="" and Request.Form("SortPath")="" then
		response.write ("<script language='javascript'>alert('��ѡ���������࣡');history.back(-1);</script>")
		response.End
	end If
	if ltrim(request.Form("ProductModel")) = "" then
		response.write ("<script language='javascript'>alert('����д��Ʒ�ͺţ�');history.back(-1);</script>")
		response.end
	end If
	if (not IsNumeric(trim(request.Form("N_Price")))) or (not IsNumeric(trim(request.Form("P_Price"))))then
		response.write ("<script language='javascript'>alert('����ȷ��д�г��۸������۸�');history.back(-1);</script>")
		response.end
	elseif trim(request.Form("N_Price"))<0 or trim(request.Form("P_Price"))<0then
		response.write ("<script language='javascript'>alert('����ȷ��д�г��۸������۸�');history.back(-1);</script>")
		response.end
	end If
	if (not IsNumeric(trim(request.Form("Stock"))))  then
		response.write ("<script language='javascript'>alert('����д/ѡ����������');history.back(-1);</script>")
		response.end
	end if
	if len(trim(Request.Form("Unit")))=0 then
		response.write ("<script language='javascript'>alert('����д/ѡ���Ʒ��λ��');history.back(-1);</script>")
		response.end
	end If
    if ltrim(request.Form("Maker")) = "" then
		response.write ("<script language='javascript'>alert('����д��Ʒ��˾��');history.back(-1);</script>")
		response.end
    end If
	if ltrim(request.Form("SmallPic")) = "" then
		response.write ("<script language='javascript'>alert('���ϴ���ƷСͼ��');history.back(-1);</script>")
		response.end
	end If
	if ltrim(request.Form("BigPic")) = "" then
		response.write ("<script language='javascript'>alert('���ϴ���Ʒ��ͼ��');history.back(-1);</script>")
		response.end
	end If
	if ltrim(request.Form("Content")) = "" then
		response.write ("<script language='javascript'>alert('����д��Ʒ��ϸ���ܣ�');history.back(-1);</script>")
		response.end
	end If
    if Result="Add" Then
	  set rsRepeat = conn.execute("select ProductNo from Qianbo_Products where ProductNo='" & trim(Request.Form("ProductNo")) & "'")
	  if not (rsRepeat.bof and rsRepeat.eof) then
		response.write "<script language='javascript'>alert('" & trim(Request.Form("ProductNo")) & "��Ʒ����Ѵ��ڣ�');history.back(-1);</script>"
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
      GroupIdName=split(Request.Form("GroupID"),"���橾")
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
					attribute1=attribute1&"����"&CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=attribute1_value&"����"&CheckStr(Request.Form("attribute"&i&"_value"),0)
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
      GroupIdName=split(Request.Form("GroupID"),"���橾")
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
					attribute1=attribute1&"����"&CheckStr(Request.Form("attribute"&i),0)
					attribute1_value=attribute1_value&"����"&CheckStr(Request.Form("attribute"&i&"_value"),0)
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
	  response.write "<script language='javascript'>alert('���óɹ�����ؾ�̬ҳ���Ѹ��£�');location.replace('ProductList.asp');</script>"
	  Else
	  response.write "<script language='javascript'>alert('���óɹ���');location.replace('ProductList.asp');</script>"
	  End If
  else
  	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Products where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
      response.write ("<center>���ݿ��¼��ȡ����</center>")
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
    response.write("δ�����")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"���橾"&rs("GroupName")&"'")
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