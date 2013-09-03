<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
call SiteInfo
if ISHTML = 1 then
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
End If
if request.QueryString("SortID")="" then
SeoTitle="产品询价"
elseif not IsNumeric(request.QueryString("SortID")) then
SeoTitle="参数错误"
elseif conn.execute("select * from Qianbo_ProductSort Where ViewFlag and  ID="&request.QueryString("SortID")).eof then
SeoTitle="参数错误"
else
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_ProductSort where ViewFlag and ID="&request.QueryString("SortID")
rs.open sql,conn,1,1
SeoTitle=rs("SortName")
rs.close
set rs=nothing
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
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
				<td style="color:#999999"><img src="images/arr6.gif" width="9" height="9" align="absmiddle" /><%=WebLocation()%> <span class="ChildSort"><% ChildSort() %></span></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="18">&nbsp;</td>
        </tr>
        <tr>
          <td><%=WebContent("Qianbo_ProductSort",request.QueryString("SortID"),"")%></td>
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
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_ProductSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
  rs.open sql,conn,1,1
  if conn.execute("select ID from Qianbo_ProductSort Where ViewFlag and ParentID=0").eof then
    response.write "<center>暂无相关信息</center>"
  end if
  do while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "ProductList.asp?SortID="&rs("ID")&""
	End If
	response.write "<a href="""&AutoLink&""">"&rs("SortName")&"</a>"
    'i=i+1
	'if i<level then call WebMenu(rs("ID"),i,level)
	'i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""ProductAdvisory.asp"" class=""agray"">产品询价</a>"&VbCrLf
  if request.QueryString("SortID")="" then
    WebLocation=WebLocation
  elseif not IsNumeric(request.QueryString("SortID")) then
    WebLocation=WebLocation&"参数错误"
  elseif conn.execute("select * from Qianbo_ProductSort Where ViewFlag and  ID="&request.QueryString("SortID")).eof then
    WebLocation=WebLocation&"参数错误"
  else
    dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_ProductSort where ViewFlag and ID="&request.QueryString("SortID")
    rs.open sql,conn,1,1
	WebLocation=WebLocation&SortPathTXT("Qianbo_ProductSort",rs("ID"))
    rs.close
    set rs=nothing
  end if
end Function

function ChildSort()
  dim ParentID
  ParentID=request.QueryString("SortID")
  if ParentID="" or (not isnumeric(ParentID)) then exit function
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_ProductSort where ViewFlag and ParentID="&ParentID&" order by ID desc"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    exit function
  else
	while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "ProductList.asp?SortID="&rs("ID")&""
	End If
	response.write "<a href="""&AutoLink&""">"&rs("SortName")&"</a> | "
	rs.movenext
	wend
  end if
  rs.close
  set rs=nothing
end function

function SortPathTXT(DataFrom,ID)
  dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "&DataFrom&" where ViewFlag and ID="&ID
  rs.open sql,conn,1,1
  if not rs.eof Then
	If ISHTML = 1 Then
		AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "ProductList.asp?SortID="&rs("ID")&""
	End If
	SortPathTXT=SortPathTXT(DataFrom,rs("ParentID"))&" - <a href="""&AutoLink&""">"&rs("SortName")&"</a>"
  end if
  rs.close
  set rs=nothing
end function

function WebContent(DataFrom,ID,SortPath)
  dim rs,sql
  dim HideSort
  set rs = server.createobject("adodb.recordset")
  if ID="" then
	SortPath="0,"
  elseif not IsNumeric(ID) then
    response.write "<center>暂无相关信息</center>"
	exit function
  elseif conn.execute("select * from "&DataFrom&" Where ViewFlag and  ID="&ID).eof then
    response.write "<center>暂无相关信息</center>"
	exit function
  else
	SortPath=conn.execute("select * from "&DataFrom&" Where ViewFlag and  ID="&ID)("SortPath")
	conn.execute("update "&DataFrom&" set ClickNumber=ClickNumber+1 Where ID="&ID)
  end if
  sql="select * from "&DataFrom&" Where not(ViewFlag) and Instr(SortPath,'"&SortPath&"')>0"
  rs.open sql,conn,1,1
  while not rs.eof
	HideSort="and not(Instr(SortPath,'"&rs("SortPath")&"')>0) "&HideSort
    rs.movenext
  wend
  rs.close
  dim idCount
  dim pages
      pages=ProInfo
  dim pagec
  dim page
      page=clng(request("Page"))
  dim pagenc
      pagenc=5
  dim pagenmax
  dim pagenmin
  dim pageprevious
  dim pagenext
  datafrom="Qianbo_Products"
  dim datawhere
	  datawhere="where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 "&HideSort& " "
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
	Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
	Response.Write "<form action=""ProductBuy.asp"" method=""Post"" name=""Inquire"" target=""_blank"">"&VbCrLf
	Response.Write "  <tr height=""32"">"&VbCrLf
	Response.Write "    <td width=""300"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;&nbsp;&nbsp;信息标题</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">产品编号</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">产品型号</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">计价单位</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">库存数量</td>"&VbCrLf
	Response.Write "    <td align=""center"" bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold"">发布日期</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
    while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&ProName&""&Separated&""&rs("ID")&"."&HTMLName&""
	Else
		AutoLink = "ProductView.asp?ID="&rs("ID")&""
	End If
	Response.Write "  <tr height=""28"">"&VbCrLf
	Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;""> <input name=""ProductNo"" type=""checkbox"" id=""ProductNo"" value="""&trim(rs("ProductNo"))&""" class=""inputnoborder""> <a href="""&AutoLink&""">"&rs("ProductName")&"</a></td>"&VbCrLf
	Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&rs("ProductNo")&"</td>"&VbCrLf
	Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&rs("ProductModel")&"</td>"&VbCrLf
	Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&rs("Unit")&"</td>"&VbCrLf
	Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&rs("Stock")&"</td>"&VbCrLf
	Response.Write "    <td align=""center"" style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">"&FormatDate(rs("Addtime"),13)&"</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
	rs.movenext
    Wend
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td height=""50"" align=""center"" colspan=""6""><input type=""image"" src=""images/buy.gif"" border=""0"" class=""inputnoborder"" /></td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
    Response.Write "</form>"&VbCrLf
	Response.Write "</table>"&VbCrLf
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
  If ISHTML = 1 Then
  If ID = "" Then
  if(page>1) then response.write ("<a href="""&AdvisoryNameDiy&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
  Else
  if(page>1) then response.write ("<a href="""&AdvisoryNameDiy&""&Separated&""&ID&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
  End If
  Else
  if(page>1) then response.write ("<a href="""& myself &"Page=1"" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
  End If
  if page-(pagenc*2+1)<=0 then
	pageprevious=1
  else
	pageprevious=page-(pagenc*2+1)
  end If
  If ISHTML = 1 Then
  If ID = "" Then
  if(pagenmin>1) then response.write ("<a href="""&AdvisoryNameDiy&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
  Else
  if(pagenmin>1) then response.write ("<a href="""&AdvisoryNameDiy&""&Separated&""&ID&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
  End If
  Else
  if(pagenmin>1) then response.write ("<a href="""& myself &"Page="& pageprevious &""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
  End If
  if(pagenmax>pagec) then pagenmax=pagec
  for i = pagenmin to pagenmax
	if(i=page) then
	  response.write ("&nbsp;<strong style=""color:red"">"& i &"</strong>&nbsp;")
	Else
	If ISHTML = 1 Then
	If ID = "" Then
		response.write ("[<a href="""&AdvisoryNameDiy&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
	Else
		response.write ("[<a href="""&AdvisoryNameDiy&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
	End If
	Else
		response.write ("[<a href="""& myself &"Page="&i&""">"& i &"</a>]")
	End If
	end if
  next
  if page+(pagenc*2+1)>=pagec then
    pagenext=pagec
  else
    pagenext=page+(pagenc*2+1)
  end If
  If ISHTML = 1 Then
  If ID = "" Then
  if(pagenmax<pagec) then response.write (" <a href="""&AdvisoryNameDiy&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""&AdvisoryNameDiy&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
  Else
  if(pagenmax<pagec) then response.write (" <a href="""&AdvisoryNameDiy&""&Separated&""&ID&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""&AdvisoryNameDiy&""&Separated&""&ID&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
  End If
  Else
  if(pagenmax<pagec) then response.write (" <a href="""& myself &"Page="& pagenext &""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""& myself &"Page="& pagec &""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
  End If
  Response.Write "    </td>"&VbCrLf
  Response.Write "  </tr>"&VbCrLf
  Response.Write "</table>"&VbCrLf
  rs.close
  set rs=nothing
end function
%>