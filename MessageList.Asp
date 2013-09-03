<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>查看留言 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
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
          <td height="20"><%=WebContent()%></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu()
  response.write "<a href=""MessageWrite.asp"">签写留言</a>"
  response.write "<a href=""MessageList.asp"">所有留言</a>"
  response.write "<a href=""MemberMessage.asp"">我的留言</a>"
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""MessageList.asp"" class=""agray"">留言反馈</a> - 查看留言"&VbCrLf
end Function

function WebContent()
  dim idCount
  dim pages
      pages=MessageInfo
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
      datafrom="Qianbo_Message "
  dim datawhere
	  datawhere="where ViewFlag "
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
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td width=""5%"" height=""25"" bgcolor=""#820000"" align=""center""><font color=""#FFFFFF""><b>"&rs("ID")&"</b></font></td>"&VbCrLf
	Response.Write "    <td width=""50%"" height=""25"" bgcolor=""#EDEEEF"" align=""left"" style=""padding-left: 8px""><font color=""#333333""><b>"&rs("Company")&"</b></font>("&rs("Linkman")&""&rs("Sex")&")</td>"&VbCrLf
	Response.Write "    <td width=""45%"" height=""25"" bgcolor=""#EDEEEF"" align=""right"" style=""padding-right: 8px""><font color=""#808080"">留言时间：</font>"&rs("Addtime")&"</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	Response.Write "</table>"&VbCrLf
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""1"">"&VbCrLf
	Response.Write "  <tr align=""left"">"&VbCrLf
	Response.Write "    <td width=""15%"" height=""20"" align=""right"" style=""padding-right: 8px"">咨询主题：</td>"&VbCrLf
	Response.Write "    <td><b>"&rs("MesName")&"</b></td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	Response.Write "  <tr align=""left"">"&VbCrLf
	Response.Write "    <td height=""20"" align=""right"" style=""padding-right: 8px"">咨询内容：</td>"&VbCrLf
	Response.Write "    <td style=""padding-right: 5px"">"&MessageContent(rs("MemID"),rs("SecretFlag"),HtmlStrReplace(rs("Content")))&"</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	If rs("ReplyContent") <> "" Then
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td>&nbsp;</td>"&VbCrLf
	Response.Write "    <td bgcolor=""#F7F7F7"" style=""line-height: 200%""><font color=""#333333""><b>管理员回复：</b></font><br />"&MessageReply (rs("MemID"),rs("MesName"),rs("SecretFlag"),rs("ReplyContent"))&"</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	End If
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td colspan=""2"" height=""5""></td>"&VbCrLf
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

function MessageContent(MemID,SecretFlag,Content)
  if SecretFlag then
    dim rs,sql
    set rs = server.createobject("adodb.recordset")
    sql="select * from Qianbo_Members where ID="&MemID
    rs.open sql,conn,1,1
    if rs("MemName")=session("MemName") and session("MemLogin")="Succeed" then
	  MessageContent= Content
	else
	  MessageContent="隐秘内容"
	end if
    rs.close
    set rs=nothing
  else
    MessageContent=Content
  end if
end function

function MessageReply(MemID,MesName,SecretFlag,ReplyContent)
  if ReplyContent<>"" then
    if SecretFlag then
      dim rs,sql
      set rs = server.createobject("adodb.recordset")
      sql="select * from Qianbo_Members where ID="&MemID
      rs.open sql,conn,1,1
      if rs("MemName")=session("MemName") and session("MemLogin")="Succeed" then
        MessageReply=MessageReply&HtmlStrReplace(ReplyContent)
	  else
        MessageReply=MessageReply&"<font color=""#CC0000"">回复已隐藏，请登录后查看。</font>"
	  end if
      rs.close
      set rs=nothing
    else
        MessageReply=MessageReply&HtmlStrReplace(ReplyContent)
    end if
  end if
end function
%>