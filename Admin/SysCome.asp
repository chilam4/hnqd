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
    <th height="22" colspan="2" sytle="line-height:150%">【系统信息】</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">服务器类型：<%=Request.ServerVariables("OS")%>(IP：<%=Request.ServerVariables("local_addr")%>)</td>
    <td width="53%" class="forumRowHighlight">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="forumRow">站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PAth")%></td>
    <td class="forumRowHighlight"> CDONTS组件：
      <%
	  On Error Resume Next
	  Server.CreateObject("CDONTS.NewMail")
	  if err=0 then 
		 response.write("<b><font color=""red"">√</font></b>")
	  else
         response.write("<b><font color=""red"">×</font></b>")
	  end if
	  err=0
      %>
      Jmail邮箱组件：
      <%
	  If Not IsObjInstalled(theInstalledObjects(13)) Then
         response.write("<b><font color=""red"">×</font></b>") 
      else
         response.write("<b><font color=""red"">√</font></b>") 
      end if
      %></td>
  </tr>
  <tr>
    <td class="forumRow">FSO文本读写：
      <%If Not testObject("scripting.filesystemobject") Then%>
      <b><font color="#FF0000">×</font></b>
      <%else%>
      <b><font color="#FF0000">√</font></b>
      <%end if%></td>
    <td class="forumRowHighlight">脚本超时时间：<%=Server.ScriptTimeout%>秒</td>
  </tr>
  <tr>
    <td class="forumRow">客户端操作系统：
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
	     vOs="类Unix"
      elseif instr(thesoft,"Mac") then
	     vOs="Mac"
      else
     	vOs="Other"
      end if
      response.Write(vOs)
      %></td>
    <td class="forumRowHighlight">返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【版本信息】</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">网站后台管理系统 <%=Str_Soft_Version%></td>
    <td width="53%" class="forumRowHighlight">程序开发：XX网络科技公司</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">管 理 员：<%=m_ManageNumber%>个 注册会员：<%=m_UserNumber%>个 留言：<%=m_Message%>(已审<%=m_MessageViewFlag%>条) 应聘信息：<%=m_Talents%>条</td>
    <td width="53%" class="forumRowHighlight">企业信息：<%=m_About%>条 下载信息：<%=m_Download%>条 友情链接：<%=m_FriendLink%>条 人才信息：<%=m_Jobs%>条</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">新闻动态：<%=m_News%>条 在线订单：<%=m_Order%>条</td>
    <td width="53%" class="forumRowHighlight">其他信息：<%=m_Others%>条 公司产品：<%=m_Products%>条</td>
  </tr>
</table>
<br />