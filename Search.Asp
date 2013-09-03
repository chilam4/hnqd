<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background:url(images/main_bg.jpg) repeat-x;">
  <tr>
    <td width="230" align="left" valign="top" style="background:url(images/left_bg.gif) repeat-y right;"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="190"><img src="images/left_nav.gif" width="190" height="28" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><div id="nav"><%=WebMenu(0,0,2)%></div></td>
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
          <td><%=WebContent(trim(request("Keyword")),request("Range"))%></td>
        </tr>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
function WebMenu(ParentID,i,level)
  if request("Range")="Void" then
	 response.write "<center>请选择搜索范围</center>"
     exit function
  end If
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_"&request("Range")&"Sort where ViewFlag and ParentID="&ParentID&" order by ID asc"
  rs.open sql,conn,1,1
  if conn.execute("select ID from Qianbo_"&request("Range")&"Sort Where ViewFlag and ParentID=0").eof then
    response.write "<center>暂无相关信息</center>"
  end if
  do while not rs.eof
	response.write "<a href="""&request("Range")&"List.asp?SortID="&rs("ID")&""">"&rs("SortName")&"</a>"
    'i=i+1
	'if i<level then call WebMenu(rs("ID"),i,level)
	'i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - 搜索 - "&VbCrLf
  if request("Range")="Void" then
	WebLocation=WebLocation&"<center>请选择搜索范围</center>"
    exit Function
  elseif trim(request("Keyword")) ="" Then
    WebLocation=WebLocation&"<font color=""red"">关键字为空</font>"
  elseif request("Range")="Product" Then
    WebLocation=WebLocation&"产品 - 关键字[<font color=""red"">"&trim(request("Keyword"))&"</font>]"
  elseif request("Range")="News" Then
    WebLocation=WebLocation&"新闻 - 关键字[<font color=""red"">"&trim(request("Keyword"))&"</font>]"
  elseif request("Range")="Down" Then
    WebLocation=WebLocation&"下载 - 关键字[<font color=""red"">"&trim(request("Keyword"))&"</font>]"
  elseif request("Range")="Others" Then
    WebLocation=WebLocation&"其他 - 关键字[<font color=""red"">"&trim(request("Keyword"))&"</font>]"
  end if
end Function

function WebContent(Keyword,Range)
  if Range="Void" then
	 response.write "<center>请选择搜索范围</center>"
     exit function
  end if
  dim idCount
  dim pages
      pages=12
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
  dim datawhere
  	  select case Range
	  case "Product"
        datafrom="Qianbo_Products"
        datawhere="where ViewFlag and ProductName like '%"&KeyWord&"%'"
	  case "News"
        datafrom="Qianbo_News"
        datawhere="where ViewFlag and NewsName like '%"&KeyWord&"%'"
	  case "Down"
        datafrom="Qianbo_Download"
        datawhere="where ViewFlag and DownName like '%"&KeyWord&"%'"
	  case "Others"
        datafrom="Qianbo_Others"
        datawhere="where ViewFlag and OthersName like '%"&KeyWord&"%'"
      end select
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
  if(idcount>0 and sqlid<>"") Then
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
	Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
	Response.Write "  <tr height=""32"">"&VbCrLf
	Response.Write "    <td width=""550"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;&nbsp;&nbsp;信息标题</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold"">发布日期</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
    while not rs.eof
  	  select case Range
	  case "Product"
		Response.Write "  <tr height=""28"">"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;<a href=""ProductView.asp?ID="&rs("ID")&""">"&rs("ProductName")&"</a></td>"&VbCrLf
		Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&FormatDate(rs("UpdateTime"),13)&"</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
	  case "News"
		Response.Write "  <tr height=""28"">"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;<a href=""NewsView.asp?ID="&rs("ID")&""">"&rs("NewsName")&"</a></td>"&VbCrLf
		Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&FormatDate(rs("Addtime"),13)&"</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
	  case "Down"
		Response.Write "  <tr height=""28"">"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;<a href=""DownView.asp?ID="&rs("ID")&""">"&rs("DownName")&"</a>("&rs("FileSize")&")</td>"&VbCrLf
		Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&FormatDate(rs("Updatetime"),13)&"</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
	  case "Others"
		Response.Write "  <tr height=""28"">"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;<a href=""OtherView.asp?ID="&rs("ID")&""">"&rs("OthersName")&"</a></td>"&VbCrLf
		Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&FormatDate(rs("UpdateTime"),13)&"</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
      end select
	  rs.movenext
    wend
	Response.Write "</table>"&VbCrLf
  else
    response.write "<center>暂无相关信息</center>"
	exit function
  end If
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