<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>我的应聘 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then
  response.Write "<script language=javascript>alert('请先登录！');location.href='Index.asp';</script>"
  response.end
end if
%>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background:url(images/main_bg.jpg) repeat-x;">
  <tr>
    <td width="230" align="left" valign="top" style="background:url(images/left_bg.gif) repeat-y right;"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="190"><img src="images/left_nav.gif" width="190" height="28" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><div id="nav"><%=WebMenu()%></div></td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
      </table>
      <!--#include file="Center_Left.asp" --></td>
    <td width="773" valign="top"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="733" style="background:url(images/bg1.gif) repeat-x left bottom;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" align="right" style="color:#999999"><img src="images/arr6.gif" width="9" height="9" align="absmiddle" /><%=WebLocation()%></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
        <tr>
          <td><%=WebContent()%></td>
        </tr>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
Function WebMenu()
If session("MemName")="" Or session("GroupID")="" Or session("MemLogin")<>"Succeed" Then
  response.write "<a href=""MemberRegister.asp"">注册会员</a>"
Else
  response.write "<a href=""MemberInfo.asp"">修改注册资料</a>"
  response.write "<a href=""MemberMessage.asp"">我的留言</a>"
  response.write "<a href=""MemberOrder.asp"">我的订单</a>"
  response.write "<a href=""MemberTalent.asp"">我的应聘</a>"
  response.write "<a href=""MemberLogin.asp?Action=Out"">退出登录</a>"
End If
End Function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MemberCenter.asp"" class=""agray"">会员中心</a> - 我的应聘"&VbCrLf
end Function

function WebContent()
  dim idCount
  dim pages
      pages=2
  dim pagec
  dim page
      page=clng(request("Page"))
  dim pagenc
      pagenc=5
  dim pagenmax
  dim pagenmin
  dim pageprevious
  dim pagenext
  dim datafrom
      datafrom="Qianbo_Talents "
  dim datawhere
	  datawhere="where MemID="&conn.execute("select * from Qianbo_Members where MemName='"&session("MemName")&"'")("ID")&" "
  dim sqlid
  dim Myself,PATH_INFO,QUERY_STRING
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" then
	    Myself = PATH_INFO & "?"
	  elseif Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself= PATH_INFO & "?" & QUERY_STRING & "&"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis
      taxis="order by id desc "
  dim i
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")
  if(idcount>0) then
    if(idcount mod pages=0)then
	  pagec=int(idcount/pages)
   	else
      pagec=int(idcount/pages)+1
    end if
    sql="select id from ["& datafrom &"] " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page
    for i=1 to rs.pagesize
	  if rs.eof then exit for
	  if(i=1)then
	    sqlid=rs("id")
	  else
	    sqlid=sqlid &","&rs("id")
	  end if
	  rs.movenext
    next
  end if
  if(idcount>0 and sqlid<>"") then
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while not rs.eof
		Response.Write "<table width=""95%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""5"">"&VbCrLf
		Response.Write "  <tr>"&VbCrLf
		Response.Write "    <td style=""background:url(images/bg1.gif) repeat-x left bottom"">您所申请的职位：<strong style=""color:red"">"&rs("TalentsName")&"</strong> 简历提交时间："&FormatDate(rs("Addtime"),13)&""&VbCrLf
		if rs("ReplyTime")<>"" then response.write "回复时间："&FormatDate(rs("ReplyTime"),13)&""
		Response.Write "    </td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
		Response.Write "  <tr>"&VbCrLf
		Response.Write "    <td style=""line-height:200%"">姓　　名："&rs("Linkman")&"<br />"&VbCrLf
		Response.Write "      性　　别："&rs("Sex")&"人<br />"&VbCrLf
		Response.Write "      出生日期："&FormatDate(rs("BirthDate"),13)&"<br />"&VbCrLf
		Response.Write "      身　　高："&int(rs("Stature"))&"CM<br />"&VbCrLf
		Response.Write "      婚姻状况："&rs("Marriage")&"<br />"&VbCrLf
		Response.Write "      户口所在地："&rs("RegResidence")&"<br />"&VbCrLf
		Response.Write "      联系地址："&rs("Address")&"<br />"&VbCrLf
		Response.Write "      邮政编码："&rs("ZipCode")&"<br />"&VbCrLf
		Response.Write "	  联系电话："&rs("Telephone")&"<br />"&VbCrLf
		Response.Write "      手机号码："&rs("Mobile")&"<br />"&VbCrLf
		Response.Write "      电子信箱："&rs("Email")&"<br />"&VbCrLf
		Response.Write "      <strong style=""color:red"">教育经历：</strong><br />"&VbCrLf
		Response.Write "      "&HtmlStrReplace(rs("EduResume"))&"<br />"&VbCrLf
		Response.Write "      <strong style=""color:red"">工作经历：</strong><br />"&VbCrLf
		Response.Write "      "&HtmlStrReplace(rs("JobResume"))&"<br />"&VbCrLf
		if rs("ReplyContent")<>"" then response.write "<strong style=""color:blue"">回复内容：</strong><br />"&HtmlStrReplace(rs("ReplyContent"))
		Response.Write "      </td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
		Response.Write "</table>"&VbCrLf
	  rs.movenext
    wend
  else
    response.write "<center>暂无相关信息</center>"
	exit function
  end if
  Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
  Response.Write "  <tr height=""35"">"&VbCrLf
  Response.Write "    <td align=""center"">"&VbCrLf
  Response.Write "共<strong style=""color:red"">"&idcount&"</strong>条记录 页次：<strong style=""color:red"">"&page&"</strong>/"&pagec&" 每页：<strong style=""color:red"">"&pages&"</strong>条记录" & vbCrLf
  pagenmin=page-pagenc
  pagenmax=page+pagenc
  if(pagenmin<1) then pagenmin=1
  if(page>1) then response.write ("<a href="""& myself &"Page=1"" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
  if page-(pagenc*2+1)<=0 then
	pageprevious=1
  else
	pageprevious=page-(pagenc*2+1)
  end if
  if(pagenmin>1) then response.write ("<a href="""& myself &"Page="& pageprevious &""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
  if(pagenmax>pagec) then pagenmax=pagec
  for i = pagenmin to pagenmax
	if(i=page) then
	  response.write ("&nbsp;<strong style=""color:red"">"& i &"</strong>&nbsp;")
	else
	  response.write ("[<a href="""& myself &"Page="& i &""">"& i &"</a>]")
	end if
  next
  if page+(pagenc*2+1)>=pagec then
    pagenext=pagec
  else
    pagenext=page+(pagenc*2+1)
  end if
  if(pagenmax<pagec) then response.write (" <a href="""& myself &"Page="& pagenext &""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""& myself &"Page="& pagec &""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
  Response.Write "    </td>"&VbCrLf
  Response.Write "  </tr>"&VbCrLf
  Response.Write "</table>"&VbCrLf
  rs.close
  set rs=nothing
end function
%>