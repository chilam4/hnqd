<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
if Instr(session("AdminPurview"),"|23,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
  response.end
end if
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,MesName,Content,ViewFlag,SecretFlag
dim mLinkman,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,AddTime
ID=request.QueryString("ID")
call MesEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="MessageEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">����ˡ��ظ����ԡ�</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">���Ա��⣺</td>
      <td width="80%" class="forumRowHighlight"><input name="MesName" type="text" id="MesName" style="width: 280" value="<%=MesName%>" maxlength="250"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�������ݣ�</td>
      <td class="forumRowHighlight"><textarea name="Content" rows="8" id="Content" style="width: 500"><%=Content%></textarea> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�����ˣ�</td>
      <td class="forumRowHighlight"><%=mLinkman%></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��λ���ƣ�</td>
      <td class="forumRowHighlight"><input name="Company" type="text" style="width: 250" value="<%=mCompany%>" maxlength="250" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">ͨ�ŵ�ַ��</td>
      <td class="forumRowHighlight"><input name="Address" type="text" style="width: 250" value="<%=mAddress%>" maxlength="250" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�������룺</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" style="width: 80" value="<%=mZipCode%>" maxlength="80" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��ϵ�绰��</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" style="width: 180" value="<%=mTelephone%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">������룺</td>
      <td class="forumRowHighlight"><input name="Fax" type="text" style="width: 180" value="<%=mFax%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�ֻ����룺</td>
      <td class="forumRowHighlight"><input name="Mobile" type="text" style="width: 180" value="<%=mMobile%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�������䣺</td>
      <td class="forumRowHighlight"><input name="Email" type="text" style="width: 180" value="<%=mEmail%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��ǰ״̬��</td>
      <td class="forumRowHighlight"><input name="SecretFlag" type="checkbox" id="SecretFlag" value="1" <%if SecretFlag then response.write ("checked")%>> ���Ļ� <input name="ViewFlag" type="checkbox" id="ViewFlag" value="1" <%if ViewFlag then response.write ("checked")%>>ͨ�����</td>
    </tr>
	<tr>
      <td align="right" class="forumRow">����ʱ�䣺</td>
      <td class="forumRowHighlight"><input name="AddTime" type="text" style="width: 180" value="<%=AddTime%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�ظ�ʱ�䣺</td>
      <td class="forumRowHighlight"><input name="ReplyTime" type="text" style="width: 180" value="<%=ReplyTime%>" maxlength="180" readonly></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�ظ����ݣ�</td>
      <td class="forumRowHighlight"><textarea name="ReplyContent" rows="8" id="ReplyContent" style="width: 500"><%=ReplyContent%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="����"> <input type="button" value="������һҳ" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
sub MesEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("MesName")))<1 then
      response.write ("<script language='javascript'>alert('����д���Ա��⣡');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("Content")))<1 then
      response.write ("<script language='javascript'>alert('����д�������ݣ�');history.back(-1);</script>")
      response.end
    end if
	if Result="Modify" then
      sql="select * from Qianbo_Message where ID="&ID
      rs.open sql,conn,1,3
      rs("MesName")=trim(Request.Form("MesName"))
      rs("Content")= StrReplace(Request.Form("Content"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SecretFlag")=1 then
        rs("SecretFlag")=Request.Form("SecretFlag")
	  else
        rs("SecretFlag")=0
	  end if
	  rs("ReplyContent")=StrReplace(Request.Form("ReplyContent"))
	  if not (trim(request.Form("ReplyContent"))="" or trim(request.Form("ReplyTime"))<>"") then
	    rs("ReplyTime")=now()
    end if
	end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('��ˡ��ظ��ɹ���');location.replace('MessageList.asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Message where ID="& ID
      rs.open sql,conn,1,1
	  MesName=rs("MesName")
	  Content=ReStrReplace(rs("Content"))
	  mLinkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  mCompany=rs("Company")
	  mAddress=rs("Address")
	  mZipCode=rs("ZipCode")
	  mTelephone=rs("Telephone")
	  mFax=rs("Fax")
	  mMobile=rs("Mobile")
	  mEmail=rs("Email")
	  ViewFlag=rs("ViewFlag")
	  SecretFlag=rs("SecretFlag")
	  AddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")
	  rs.close
      set rs=nothing
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  else
    GuestInfo="<font color='green'>��Ա��</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"'>"&Guest&"</a>"&Sex
  end if
  rs.close
  set rs=nothing
end function
%>