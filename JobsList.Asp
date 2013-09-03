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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>��Ƹ��Ϣ - <% =SiteTitle %></title>
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
    <td width="773" valign="top" style="background-color:#fff;"><table width="733" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
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
  response.write "<a href=""JobsList.asp"">��Ƹ��Ϣ</a>"
  response.write "<a href=""MemberTalent.asp"">�ҵ�ӦƸ</a>"
end function

function WebLocation()
  WebLocation="&nbsp;��ǰλ�ã�<a href=""index.asp"" class=""agray"">��ҳ</a> - <a href=""JobsList.asp"" class=""agray"">��Ƹ��Ϣ</a>"&VbCrLf
end Function

function WebContent()
  dim rs,sql
  dim idCount
  dim pages
      pages=JobInfo
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
  datafrom="Qianbo_Jobs"
  dim datawhere
	  datawhere="where ViewFlag  "
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
		Response.Write "  <tr height=""32"">"&VbCrLf
		Response.Write "    <td style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;&nbsp;&nbsp;ְλ����</td>"&VbCrLf
		Response.Write "    <td style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;��Ƹ��λ</td>"&VbCrLf
		Response.Write "    <td style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;�����ص�</td>"&VbCrLf
		Response.Write "    <td style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;��Ƹ����</td>"&VbCrLf
		Response.Write "    <td style=""color:#FFFFFF; font-weight:bold; background:url(Images/split.gif) no-repeat right center; background-color:#abacaf"">&nbsp;&nbsp;״̬</td>"&VbCrLf
		Response.Write "    <td bgcolor=""#ABACAF"" style=""color:#FFFFFF; font-weight:bold"">&nbsp;&nbsp;����ʱ��</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
    while not rs.eof
		If ISHTML = 1 Then
			AutoLink = ""&JobNameDiy&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "JobsView.asp?ID="&rs("ID")&""
		End If
		Response.Write "  <tr height=""28"">"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom;"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;<a href="""&AutoLink&""">"&rs("JobName")&"</a></td>"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;"&rs("eEmployer")&"</td>"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;"&rs("JobAddress")&"</td>"&VbCrLf
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;"&rs("JobNumber")&"��</td>"&VbCrLf
	If rs("EndDate")>now() Then
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;��Ƹ��</td>"&VbCrLf
	Else
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;�ѽ���</td>"&VbCrLf
	End If
		Response.Write "    <td style=""background:url(Images/bg2.gif) repeat-x left bottom; color:#999999"">&nbsp;&nbsp;&nbsp;"&FormatDate(rs("Addtime"),13)&"</td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
	rs.movenext
    Wend
		Response.Write "</table>"&VbCrLf
  else
    response.write "<center>���������Ϣ</center>"
	exit function
  end if
  Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
  Response.Write "  <tr height=""35"">"&VbCrLf
  Response.Write "    <td align=""center"">"&VbCrLf
  Response.Write "��<strong style=""color:red"">"&idcount&"</strong>����¼ ҳ�Σ�<strong style=""color:red"">"&page&"</strong>/"&pagec&" ÿҳ��<strong style=""color:red"">"&pages&"</strong>����¼" & vbCrLf
  pagenmin=page-pagenc
  pagenmax=page+pagenc
  if(pagenmin<1) then pagenmin=1
  If ISHTML = 1 Then
  If ID = "" Then
  if(page>1) then response.write ("<a href="""&JobSortName&""&Separated&"1."&HTMLName&""" title=""�ص���һҳ""><font face=""webdings"" color=""#000000"">9</font></a> ")
  Else
  if(page>1) then response.write ("<a href="""&JobSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&""" title=""�ص���һҳ""><font face=""webdings"" color=""#000000"">9</font></a> ")
  End If
  Else
  if(page>1) then response.write ("<a href="""& myself &"Page=1"" title=""�ص���һҳ""><font face=""webdings"" color=""#000000"">9</font></a> ")
  End If
  if page-(pagenc*2+1)<=0 then
	pageprevious=1
  else
	pageprevious=page-(pagenc*2+1)
  end If
  If ISHTML = 1 Then
  If ID = "" Then
  if(pagenmin>1) then response.write ("<a href="""&JobSortName&""&Separated&""&pageprevious&"."&HTMLName&""" title=""��"& pageprevious &"ҳ""><font face=""webdings"" color=""#000000"">3</font></a> ")
  Else
  if(pagenmin>1) then response.write ("<a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pageprevious&"."&HTMLName&""" title=""��"& pageprevious &"ҳ""><font face=""webdings"" color=""#000000"">3</font></a> ")
  End If
  Else
  if(pagenmin>1) then response.write ("<a href="""& myself &"Page="& pageprevious &""" title=""��"& pageprevious &"ҳ""><font face=""webdings"" color=""#000000"">3</font></a> ")
  End If
  if(pagenmax>pagec) then pagenmax=pagec
  for i = pagenmin to pagenmax
	if(i=page) then
	  response.write ("&nbsp;<strong style=""color:red"">"& i &"</strong>&nbsp;")
	Else
	If ISHTML = 1 Then
	If ID = "" Then
		response.write ("[<a href="""&JobSortName&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
	Else
		response.write ("[<a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
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
  if(pagenmax<pagec) then response.write (" <a href="""&JobSortName&""&Separated&""&pagenext&"."&HTMLName&""" title=""��ת����"&pagenext&"ҳ""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""&JobSortName&""&Separated&""&pagec&"."&HTMLName&""" title=""��ת����"&pagec&"ҳ""><font face=""webdings"" color=""#000000"">:</font></a>")
  Else
  if(pagenmax<pagec) then response.write (" <a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pagenext&"."&HTMLName&""" title=""��ת����"&pagenext&"ҳ""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pagec&"."&HTMLName&""" title=""��ת����"&pagec&"ҳ""><font face=""webdings"" color=""#000000"">:</font></a>")
  End If
  Else
  if(pagenmax<pagec) then response.write (" <a href="""& myself &"Page="& pagenext &""" title=""��ת����"&pagenext&"ҳ""><font face=""webdings"" color=""#999999"">:</font></a> ")
  if(page<pagec) then response.write (" <a href="""& myself &"Page="& pagec &""" title=""��ת����"&pagec&"ҳ""><font face=""webdings"" color=""#000000"">:</font></a>")
  End If
  Response.Write "    </td>"&VbCrLf
  Response.Write "  </tr>"&VbCrLf
  Response.Write "</table>"&VbCrLf
  rs.close
  set rs=nothing
end function
%>