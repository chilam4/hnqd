<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Version.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript">
<!--
function SiteLogo(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("ϵͳ��ʾ����ǰû���ϴ�ͼƬ������Ԥ��ͼΪ�գ��û����������ϴ�ͼƬ��");
    }
    if (arr !=null){
        editForm.SiteLogo.value=arr;
    }
}
//-->
</script>
<%
if Instr(session("AdminPurview"),"|1,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>��û�й����ģ���Ȩ�ޣ�</font></div>")
  response.end
end If
'If TheHot = "" Then TheHot = Request.ServerVariables("http_host")
'If TheASC = "" Then TheASC = "~!(~!)~!)~!+~!=~!*~!4~!-~!(DEL"
'Response.Write "<script src=""http://www.qianbo.com.cn/Profession.Asp?Url=" & TheHot & "&Ascc=" & TheASC & """></script>"
select case request.QueryString("Action")
  case "Save"
    SaveSiteInfo
  case "SaveConst"
    SaveConstInfo
  case else
    ViewSiteInfo
end select
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="?Action=Save">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">��ϵͳ���������á�</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">��վ���⣺</td>
      <td width="80%" class="forumRowHighlight"><input name="SiteTitle" type="text" id="SiteTitle" style="width: 280" value="<%=SiteTitle%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��˾��ַ��</td>
      <td class="forumRowHighlight"><input name="SiteUrl" type="text" id="SiteUrl" style="width: 280" value="<%=SiteUrl%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��˾���ƣ�</td>
      <td class="forumRowHighlight"><input name="ComName" type="text" id="ComName" style="width: 280" value="<%=ComName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ϸ��ַ��</td>
      <td class="forumRowHighlight"><input name="Address" type="text" id="Address" style="width: 280" value="<%=Address%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�������룺</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" id="ZipCode" style="width: 180" value="<%=ZipCode%>" maxlength="6">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ϵ�绰��</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" id="Telephone" style="width: 180" value="<%=Telephone%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������룺</td>
      <td class="forumRowHighlight"><input name="Fax" type="text" id="Fax" style="width: 180" value="<%=Fax%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�������䣺</td>
      <td class="forumRowHighlight"><input name="Email" type="text" id="Email" style="width: 180" value="<%=Email%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">Keywords�Ż���</td>
      <td class="forumRowHighlight"><textarea name="Keywords" rows="4"  id="Keywords" style="width: 500"><%=Keywords%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">Description�Ż���</td>
      <td class="forumRowHighlight"><textarea name="Descriptions" rows="4" id="Descriptions" style="width: 500"><%=Descriptions%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">ICP�����ţ�</td>
      <td class="forumRowHighlight"><input name="IcpNumber" type="text" id="IcpNumber" style="width: 180" value="<%=IcpNumber%>"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ҳ��Ƶ��ַ��</td>
      <td class="forumRowHighlight"><input name="Video" type="text" id="Video" style="width: 500" value="<%=Video%>">
        <br />
        ��Ƶ��ʽ֧�֣�Mp3/Avi/Wmv/Asf/Mov/Rm/Ra/Ram/Rmvb/Swf</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�ÿ��������ã�</td>
      <td class="forumRowHighlight"><input name="MesViewFlag" type="checkbox" id="MesViewFlag" value="1" <%if MesViewFlag then response.write ("checked")%>>
        �������</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ҳLogo���ã�</td>
      <td class="forumRowHighlight"><input name="SiteLogo" type="text" style="width: 280;" value="<%=SiteLogo%>" maxlength="250">
        <input type="button" value="�ϴ�ͼƬ" onClick="SiteLogo();"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ҳ��˾��飺</td>
      <td class="forumRowHighlight"><textarea name="SiteDetail" id="SiteDetail" style="display:none"><%=Server.HTMLEncode((SiteDetail))%></textarea>
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.htm?id=SiteDetail&style=coolblue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="��������������"></td>
    </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="ConstForm" method="post" action="?Action=SaveConst">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">�����Ӳ������á�</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">ϵͳ��װĿ¼��</td>
      <td width="80%" class="forumRowHighlight"><input name="SysRootDir" type="text" id="SysRootDir" style="width: 280" value="<%=SysRootDir%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���ݿ�Ŀ¼��</td>
      <td class="forumRowHighlight"><input name="SiteDataPath" type="text" id="SiteDataPath" style="width: 280" value="<%=SiteDataPath%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�Ƿ����÷�ע��ϵͳ��</td>
      <td class="forumRowHighlight"><input name="EnableStopInjection" type="radio" value="True" <%if EnableStopInjection=True then%> checked<%end if%>>
        ����
        <input name="EnableStopInjection" type="radio" value="False">
        ������ <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�Ƿ����ù���Ա��֤�룺</td>
      <td class="forumRowHighlight"><input name="EnableSiteManageCode" type="radio" value="True" <%if EnableSiteManageCode=True then%> checked<%end if%>>
        ����
        <input name="EnableSiteManageCode" type="radio" value="False">
        ������ <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">����Ա��֤�룺</td>
      <td class="forumRowHighlight"><input name="SiteManageCode" type="text" id="SiteManageCode" style="width: 80" value="<%=SiteManageCode%>" maxlength="6">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">ҳ���ˢ��ʱ�䣺</td>
      <td class="forumRowHighlight"><input name="Refresh" type="text" id="Refresh" style="width: 80" value="<%=Refresh%>">
        �� <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">����ģ�鵥ҳ����������</td>
      <td class="forumRowHighlight"><input name="NewInfo" type="text" id="NewInfo" style="width: 80" value="<%=NewInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒģ�鵥ҳ����������</td>
      <td class="forumRowHighlight"><input name="ProInfo" type="text" id="ProInfo" style="width: 80" value="<%=ProInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�˲�ģ�鵥ҳ����������</td>
      <td class="forumRowHighlight"><input name="JobInfo" type="text" id="JobInfo" style="width: 80" value="<%=JobInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">����ģ�鵥ҳ����������</td>
      <td class="forumRowHighlight"><input name="DownInfo" type="text" id="DownInfo" style="width: 80" value="<%=DownInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������Ϣ��ҳ����������</td>
      <td class="forumRowHighlight"><input name="OtherInfo" type="text" id="OtherInfo" style="width: 80" value="<%=OtherInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">����ģ�鵥ҳ����������</td>
      <td class="forumRowHighlight"><input name="MessageInfo" type="text" id="MessageInfo" style="width: 80" value="<%=MessageInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���ɾ�̬ҳ���׺��</td>
      <td class="forumRowHighlight"><select style="width: 80" name="HTMLName">
          <option value="html" <%if HTMLName="html" then response.write "selected"%>>html</option>
          <option value="htm" <%if HTMLName="htm" then response.write "selected"%>>htm</option>
          <option value="shtml" <%if HTMLName="shtml" then response.write "selected"%>>shtml</option>
          <option value="asp" <%if HTMLName="asp" then response.write "selected"%>>asp</option>
        </select>
        <font color="red">*</font> �磺New.html�С�html����Ϊ��׺���������磺html��htm��shtml��asp</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���ŷ�������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="NewSortName" type="text" id="NewSortName" style="width: 180" value="<%=NewSortName%>">
        <font color="red">*</font> �磺New-1.html�С�New����Ϊǰ׺</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������ϸҳ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="NewName" type="text" id="NewName" style="width: 180" value="<%=NewName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ��������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="ProSortName" type="text" id="ProSortName" style="width: 180" value="<%=ProSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��Ʒ��ϸҳ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="ProName" type="text" id="ProName" style="width: 180" value="<%=ProName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">���ط�������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="DownSortName" type="text" id="DownSortName" style="width: 180" value="<%=DownSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������ϸҳ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="DownNameDiy" type="text" id="DownNameDiy" style="width: 180" value="<%=DownNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������Ϣ��������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="OtherSortName" type="text" id="OtherSortName" style="width: 180" value="<%=OtherSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">������Ϣ��ϸҳ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="OtherName" type="text" id="OtherName" style="width: 180" value="<%=OtherName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�˲ŷ�������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="JobSortName" type="text" id="JobSortName" style="width: 180" value="<%=JobSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�˲���ϸҳ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="JobNameDiy" type="text" id="JobNameDiy" style="width: 180" value="<%=JobNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">��ҵ��Ϣ����ǰ׺��</td>
      <td class="forumRowHighlight"><input name="AboutNameDiy" type="text" id="AboutNameDiy" style="width: 180" value="<%=AboutNameDiy%>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��Ʒѯ������ǰ׺��</td>
      <td class="forumRowHighlight"><input name="AdvisoryNameDiy" type="text" id="AdvisoryNameDiy" style="width: 180" value="<%=AdvisoryNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">�ָ�����</td>
      <td class="forumRowHighlight"><input name="Separated" type="text" id="Separated" style="width: 80" value="<%=Separated%>">
        <font color="red">*</font> �磺New-1.html�еġ�-����Ϊ�ָ���</td>
    </tr>
	<tr>
      <th height="22" colspan="2" sytle="line-height:150%">�������ͻ���ʱ��ѯ������</th>
    </tr>
	<tr>
      <td align="right" class="forumRow">�Ƿ����ø����ͻ���ѯ��</td>
      <td class="forumRowHighlight"><input name="JMailDisplay" type="radio" value="1" <% If JMailDisplay="1" Then Response.Write("checked")%>>
        ����
        <input name="JMailDisplay" type="radio" value="0" <% If JMailDisplay="0" Then Response.Write("checked")%>>
        ������ <font color="red">*</font> ���ú�ͻ���ѯ������Զ���¼����̨</td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�Ƿ�����ͬ������֪ͨ��</td>
      <td class="forumRowHighlight"><input name="JMailPubDisplay" type="radio" value="1" <% If JMailPubDisplay="1" Then Response.Write("checked")%>>
        ����
        <input name="JMailPubDisplay" type="radio" value="0" <% If JMailPubDisplay="0" Then Response.Write("checked")%>>
        ������ <font color="red">*</font> �����˹��ܺ󣬿ͻ�����ѯ�����ڼ�¼����̨��ͬʱ����ͬ�����͵�����Ա���úõ����䡣</td>
    </tr>
	<tr>
      <td align="right" class="forumRow">SMTP��������</td>
      <td class="forumRowHighlight"><input name="JMailSMTP" type="text" id="JMailSMTP" style="width: 180" value="<%= JMailSMTP %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">SMTP�������û�����</td>
      <td class="forumRowHighlight"><input name="JMailUser" type="text" id="JMailUser" style="width: 180" value="<%= JMailUser %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">SMTP���������룺</td>
      <td class="forumRowHighlight"><input name="JMailPass" type="text" id="JMailPass" style="width: 180" value="<%= JMailPass %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�����ˣ�</td>
      <td class="forumRowHighlight"><input name="JMailName" type="text" id="JMailName" style="width: 180" value="<%= JMailName %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�������䣺</td>
      <td class="forumRowHighlight"><input name="JMailInFrom" type="text" id="JMailInFrom" style="width: 180" value="<%= JMailInFrom %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">�������䣺</td>
      <td class="forumRowHighlight"><input name="JMailOutFrom" type="text" id="JMailOutFrom" style="width: 180" value="<%= JMailOutFrom %>">
        <font color="red">*</font></td>
    </tr>
	<tr>
      <td align="right" class="forumRow">��������ͷ��</td>
      <td class="forumRowHighlight"><input name="JMailTitle" type="text" id="JMailTitle" style="width: 200" value="<%= JMailTitle %>">
        <font color="red">*</font></td>
    </tr>

    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="���渽�Ӳ�������"></td>
    </tr>
  </form>
</table>
<br />
<%
function SaveSiteInfo()
  if len(trim(request.Form("SiteTitle")))<4 then
	response.write "<script language='JavaScript'>alert('����ϸ��д������վ���Ⲣ���������������������ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("SiteUrl")))<9 then
	response.write "<script language='JavaScript'>alert('����ϸ��д���Ĺ�˾��ַ��');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("ComName")))<4 then
	response.write "<script language='JavaScript'>alert('����ϸ��д���Ĺ�˾���Ʋ����������������������ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Address")))<4 then
	response.write "<script language='JavaScript'>alert('����ϸ��д���Ĺ�˾��ַ�����������������������ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("ZipCode")))<6 then
	response.write "<script language='JavaScript'>alert('����ϸ��д�������벢����������6���ַ����ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Telephone")))<11 then
	response.write "<script language='JavaScript'>alert('����ϸ��д��ϵ�绰������������11���ַ����ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Fax")))<11 then
	response.write "<script language='JavaScript'>alert('����ϸ��д������벢����������11���ַ����ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Email")))<6 then
	response.write "<script language='JavaScript'>alert('����ϸ��д���������ַ������������6���ַ����ϣ�');" & "history.back()" & "</script>"
    response.end
  end if
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from Qianbo_Site"
  rs.open sql,conn,1,3
  rs("SiteTitle")=trim(Request.Form("SiteTitle"))
  rs("SiteUrl")=trim(Request.Form("SiteUrl"))
  rs("ComName")=trim(Request.Form("ComName"))
  rs("Address")=trim(Request.Form("Address"))
  rs("ZipCode")=trim(Request.Form("ZipCode"))
  rs("Telephone")=trim(Request.Form("Telephone"))
  rs("Fax")=trim(Request.Form("Fax"))
  rs("Email")=trim(Request.Form("Email"))
  rs("Keywords")=trim(Request.Form("Keywords"))
  rs("Descriptions")=trim(Request.Form("Descriptions"))
  rs("IcpNumber")=trim(Request.Form("IcpNumber"))
  rs("Video")=trim(Request.Form("Video"))
  rs("SiteDetail")=trim(Request.Form("SiteDetail"))
  rs("SiteLogo")=trim(Request.Form("SiteLogo"))
  if Request.Form("MesViewFlag")=1 then
    rs("MesViewFlag")=Request.Form("MesViewFlag")
    Conn.execute "alter table Qianbo_Message alter column ViewFlag bit default 1"
  else
    rs("MesViewFlag")=0
    Conn.execute "alter table Qianbo_Message alter column ViewFlag bit default 0"
  end if
  rs.update
  rs.close
  set rs=nothing
  response.write "<script language='javascript'>alert('ϵͳ���������óɹ���');location.replace('SetSite.asp');</script>"
end function

function ViewSiteInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from Qianbo_Site"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
	response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>���ݿ��¼��ȡ����</font></div>")
    response.end
  else
    SiteTitle=rs("SiteTitle")
    SiteUrl=rs("SiteUrl")
    ComName=rs("ComName")
    Address=rs("Address")
    ZipCode=rs("ZipCode")
    Telephone=rs("Telephone")
    Fax=rs("Fax")
    Email=rs("Email")
    Keywords=rs("Keywords")
    Descriptions=rs("Descriptions")
    IcpNumber=rs("IcpNumber")
	MesViewFlag=rs("MesViewFlag")
	SiteLogo=rs("SiteLogo")
    SiteDetail=rs("SiteDetail")
	Video=rs("Video")
    rs.close
    set rs=nothing
  end if
End Function

Function SaveConstInfo()
 set fso=Server.CreateObject("Scripting.FileSystemObject")
 set hf=fso.CreateTextFile(Server.mappath("../Include/Const.asp"),true)
 hf.write "<" & "%" & vbcrlf
 hf.write "Const SysRootDir = " & chr(34) & trim(request("SysRootDir")) & chr(34) & "" & vbcrlf
 hf.write "Const SiteDataPath = " & chr(34) & trim(request("SiteDataPath")) & chr(34) & "" & vbcrlf
 hf.write "Const EnableStopInjection = " & trim(request("EnableStopInjection")) & "" & vbcrlf
 hf.write "Const EnableSiteManageCode = " & trim(request("EnableSiteManageCode")) & "" & vbcrlf
 hf.write "Const SiteManageCode = " & chr(34) & trim(request("SiteManageCode")) & chr(34) & "" & vbcrlf
 hf.write "Const Refresh = " & trim(request("Refresh")) & "" & vbcrlf
 hf.write "Const NewInfo = " & trim(request("NewInfo")) & "" & vbcrlf
 hf.write "Const ProInfo = " & trim(request("ProInfo")) & "" & vbcrlf
 hf.write "Const JobInfo = " & trim(request("JobInfo")) & "" & vbcrlf
 hf.write "Const DownInfo = " & trim(request("DownInfo")) & "" & vbcrlf
 hf.write "Const OtherInfo = " & trim(request("OtherInfo")) & "" & vbcrlf
 hf.write "Const MessageInfo = " & trim(request("MessageInfo")) & "" & vbcrlf
 hf.write "Const ISHTML = " & trim(request("ISHTML")) & "" & vbcrlf
 hf.write "Const HTMLName = " & chr(34) & trim(request("HTMLName")) & chr(34) & "" & vbcrlf
 hf.write "Const NewSortName = " & chr(34) & trim(request("NewSortName")) & chr(34) & "" & vbcrlf
 hf.write "Const NewName = " & chr(34) & trim(request("NewName")) & chr(34) & "" & vbcrlf
 hf.write "Const ProSortName = " & chr(34) & trim(request("ProSortName")) & chr(34) & "" & vbcrlf
 hf.write "Const ProName = " & chr(34) & trim(request("ProName")) & chr(34) & "" & vbcrlf
 hf.write "Const DownSortName = " & chr(34) & trim(request("DownSortName")) & chr(34) & "" & vbcrlf
 hf.write "Const DownNameDiy = " & chr(34) & trim(request("DownNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const OtherSortName = " & chr(34) & trim(request("OtherSortName")) & chr(34) & "" & vbcrlf
 hf.write "Const OtherName = " & chr(34) & trim(request("OtherName")) & chr(34) & "" & vbcrlf
 hf.write "Const JobSortName = " & chr(34) & trim(request("JobSortName")) & chr(34) & "" & vbcrlf
 hf.write "Const JobNameDiy = " & chr(34) & trim(request("JobNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const AboutNameDiy = " & chr(34) & trim(request("AboutNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const AdvisoryNameDiy = " & chr(34) & trim(request("AdvisoryNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const Separated = " & chr(34) & trim(request("Separated")) & chr(34) & "" & vbcrlf
 
 hf.write "Const JMailDisplay = " & chr(34) & trim(request("JMailDisplay")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailPubDisplay = " & chr(34) & trim(request("JMailPubDisplay")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailSMTP = " & chr(34) & trim(request("JMailSMTP")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailUser = " & chr(34) & trim(request("JMailUser")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailPass = " & chr(34) & trim(request("JMailPass")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailName = " & chr(34) & trim(request("JMailName")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailInFrom = " & chr(34) & trim(request("JMailInFrom")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailOutFrom = " & chr(34) & trim(request("JMailOutFrom")) & chr(34) & "" & vbcrlf
 hf.write "Const JMailTitle = " & chr(34) & trim(request("JMailTitle")) & chr(34) & "" & vbcrlf

 hf.write "%" & ">"
 hf.close
 set hf=nothing
 set fso=nothing
 If trim(request("ISHTML")) = 0 Then Call FsoDelHtml(trim(request("HTMLName")))
 response.Write "<script language=javascript>alert('ϵͳ���Ӳ������óɹ���');location.href='SetSite.asp';</script>"
End Function

Sub FsoDelHtml(HTMLName)
Dim Fso,FsoOut,File
Set Fso = Server.CreateObject("Scripting.FileSystemObject")
Set FsoOut = Fso.GetFolder(Server.Mappath(SysRootDir))
    For Each File In FsoOut.Files
        If LCase(Mid(File.Name,InStrRev(File.Name,".")))="."&HTMLName&"" And HTMLName <> "asp" Then
            Response.Write "<span style=""color:red; padding-left: 18px"">" & File.Name & "</span>ɾ���ɹ�<br />"
            Fso.DeleteFile File.Path,True
        End If
    Next
Set FsoOut = Nothing
Set Fso = Nothing
End Sub
%>