<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="../Include/Version.asp" -->
<%
m_SQL = "select count(*) from Qianbo_Admin"
set rs = conn.Execute(m_SQL)
m_ManageNumber = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Members"
set rs = conn.Execute(m_SQL)
m_UserNumber = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Message"
set rs = conn.Execute(m_SQL)
m_Message = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Message where ViewFlag = 1"
set rs = conn.Execute(m_SQL)
m_MessageViewFlag = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_About"
set rs = conn.Execute(m_SQL)
m_About = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Download"
set rs = conn.Execute(m_SQL)
m_Download = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_FriendLink"
set rs = conn.Execute(m_SQL)
m_FriendLink = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Jobs"
set rs = conn.Execute(m_SQL)
m_Jobs = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_News"
set rs = conn.Execute(m_SQL)
m_News = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Order"
set rs = conn.Execute(m_SQL)
m_Order = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Others"
set rs = conn.Execute(m_SQL)
m_Others = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Products"
set rs = conn.Execute(m_SQL)
m_Products = rs(0)
rs.Close
set rs=Nothing

m_SQL = "select count(*) from Qianbo_Talents"
set rs = conn.Execute(m_SQL)
m_Talents = rs(0)
rs.Close
set rs=nothing
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/Admin_style.css">
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">��ϵͳ��Ϣ��</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">���������ͣ�<%=Request.ServerVariables("OS")%>(IP��<%=Request.ServerVariables("local_addr")%>)</td>
    <td width="53%" class="forumRowHighlight">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="forumRow">վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PAth")%></td>
    <td class="forumRowHighlight"> CDONTS�����
      <%
	  On Error Resume Next
	  Server.CreateObject("CDONTS.NewMail")
	  if err=0 then 
		 response.write("<b><font color=""red"">��</font></b>")
	  else
         response.write("<b><font color=""red"">��</font></b>")
	  end if
	  err=0
      %>
      Jmail���������
      <%
	  If Not IsObjInstalled(theInstalledObjects(13)) Then
         response.write("<b><font color=""red"">��</font></b>") 
      else
         response.write("<b><font color=""red"">��</font></b>") 
      end if
      %></td>
  </tr>
  <tr>
    <td class="forumRow">FSO�ı���д��
      <%If Not testObject("scripting.filesystemobject") Then%>
      <b><font color="#FF0000">��</font></b>
      <%else%>
      <b><font color="#FF0000">��</font></b>
      <%end if%></td>
    <td class="forumRowHighlight">�ű���ʱʱ�䣺<%=Server.ScriptTimeout%>��</td>
  </tr>
  <tr>
    <td class="forumRow">�ͻ��˲���ϵͳ��
      <%
      dim thesoft,vOS
      thesoft=Request.ServerVariables("HTTP_USER_AGENT")
      if instr(thesoft,"Windows NT 5.0") then
	     vOS="Microsoft Windows 2000"
      elseif instr(thesoft,"Windows NT 5.2") then
	     vOs="Microsoft Windows 2003"
      elseif instr(thesoft,"Windows NT 5.1") then
         vOs="Microsoft Windows XP"
      elseif instr(thesoft,"Windows NT") then
       	 vOs="Microsoft Windows NT"
      elseif instr(thesoft,"Windows 9") then
	     vOs="Microsoft Windows 9x"
      elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	     vOs="��Unix"
      elseif instr(thesoft,"Mac") then
	     vOs="Mac"
      else
     	vOs="Other"
      end if
      response.Write(vOs)
      %></td>
    <td class="forumRowHighlight">���ط�������������Ķ˿ڣ�<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">���汾��Ϣ��</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">��վ��̨����ϵͳ <%=Str_Soft_Version%></td>
    <td width="53%" class="forumRowHighlight">���򿪷���XX����Ƽ���˾</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">�� �� Ա��<%=m_ManageNumber%>�� ע���Ա��<%=m_UserNumber%>�� ���ԣ�<%=m_Message%>(����<%=m_MessageViewFlag%>��) ӦƸ��Ϣ��<%=m_Talents%>��</td>
    <td width="53%" class="forumRowHighlight">��ҵ��Ϣ��<%=m_About%>�� ������Ϣ��<%=m_Download%>�� �������ӣ�<%=m_FriendLink%>�� �˲���Ϣ��<%=m_Jobs%>��</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">���Ŷ�̬��<%=m_News%>�� ���߶�����<%=m_Order%>��</td>
    <td width="53%" class="forumRowHighlight">������Ϣ��<%=m_Others%>�� ��˾��Ʒ��<%=m_Products%>��</td>
  </tr>
</table>
<br />