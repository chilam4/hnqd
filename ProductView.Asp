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
ID=request.QueryString("ID")
if ID <> "" or IsNumeric(ID) then
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Products where ViewFlag and ID="&ID
rs.open sql,conn,1,3
If rs("SeoKeywords") <> "" Then
SeoKeywords=rs("SeoKeywords")
Else
SeoKeywords=rs("ProductName")
End If
If rs("SeoDescription") <> "" Then
SeoDescription=rs("SeoDescription")
Else
SeoDescription=rs("ProductName")
End If
SeoTitle=rs("ProductName")
End If
rs.close
set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
<meta name="keywords" content="<% =SeoKeywords %>" />
<meta name="description" content="<% =SeoDescription %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
<link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
<script src="Scripts/prototype.js" type="text/javascript"></script>
<script src="Scripts/scriptaculous.js?load=effects,builder" type="text/javascript"></script>
<script src="Scripts/lightbox.js" type="text/javascript"></script>
</head>
<!--#include file="Top.asp" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:1px; background:url(images/main_bg.jpg) repeat-x;">
  <tr>
    <td width="230" align="left" valign="top" style="background:url(images/left_bg.gif) repeat-y right;"><table width="190" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:19px;">
        <tr>
          <td width="190"><img src="images/left_nav.gif" width="190" height="28" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><%Folder(0)%></td>
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
          <td><table width="733" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="30" height="36"><img src="images/title_bg1.gif" width="30" height="36" /></td>
                <td width="673" align="center" background="images/title_bg_center.jpg" style="color:#000000; font-size:14px; font-weight:bold">
<%
  ID=request.QueryString("ID")
  if ID="" or not IsNumeric(ID) then
    response.write "<center>暂无相关信息</center>"
  elseif conn.execute("select * from Qianbo_Products Where ViewFlag and  ID="&ID).eof then
    response.write "<center>暂无相关信息</center>"
  else
	dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_Products where ViewFlag and ID="&ID
    rs.open sql,conn,1,3
	Response.Write ""&rs("ProductName")&""&VbCrLf
  End If
%></td>
                <td width="30"><img src="images/title_bg2.gif" width="30" height="36" /></td>
              </tr>
              <tr>
                <td><img src="images/title_bg3.gif" width="30" height="14" /></td>
                <td background="images/title_bg_center2.jpg">&nbsp;</td>
                <td><img src="images/title_bg4.gif" width="30" height="14" /></td>
              </tr>
            </table></td>
        </tr>
<%
if ViewNoRight(rs("GroupID"),rs("Exclusive")) Then
Response.Write "  <tr>"&VbCrLf
Response.Write "    <td>"&VbCrLf
Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">"&VbCrLf
Response.Write "  <tr height=""25"">"&VbCrLf
Response.Write "    <td>产品名称："&rs("ProductName")&"<br><br>"&VbCrLf
Response.Write "    产品编号："&rs("ProductNo") &"<br><br>"&VbCrLf
Response.Write "    产品型号："&rs("ProductModel") &"<br><br>"&VbCrLf
Response.Write "    市场价格："&rs("N_Price")&"元/"&rs("Unit")&"<br><br>"&VbCrLf
Response.Write "    批发价格："&rs("P_Price")&"元/"&rs("Unit")&"<br><br>"&VbCrLf
Response.Write "    更新时间："&FormatDate(rs("Addtime"),13)&"<br><br>"&VbCrLf
Response.Write "    出品单位："&rs("Maker") &"</td>"&VbCrLf
Response.Write "    <td style=""width:260px; text-align:center""><a href="""&rs("BigPic")&""" title="""&rs("ProductName")&""" rel=""lightbox"" target=""_blank""><img src="""&rs("BigPic")&""" title="""&rs("ProductName")&""" style=""border:1px #ccc;"" onload=""javascript:DrawImage(this,260,260);"" name=""rImage"" title="""&rs("ProductName")&""" border=""0""></a><br /><br /><a href=""ProductBuy.asp?ProductNo="&rs("ProductNo")&""" title=""立即订购："&rs("ProductName")&"""><img src=""Images/buy.gif"" border=""0"" /></a></td>"&VbCrLf
Response.Write "  </tr>"&VbCrLf
attribute1=rs("attribute1")
attribute1_value=rs("attribute1_value")
if attribute1 <> "" and attribute1_value <> "" then
	attribute1_1=Split(attribute1,"§§§")
	attribute1_value_1=Split(attribute1_value,"§§§")
for i=0 to ubound(attribute1_value_1)
Response.Write "  <tr height=""25"">"&VbCrLf
Response.Write "    <td>"&attribute1_1(i)&"："&attribute1_value_1(i)&"</td>"&VbCrLf
Response.Write "  </tr>"&VbCrLf
Next
End If
Response.Write "</table>"&VbCrLf
Response.Write "    </td>"&VbCrLf
Response.Write "  </tr>"&VbCrLf
Response.Write "  <tr>"&VbCrLf
Response.Write "    <td height=""33"" style=""background:url(Images/bg2.gif) repeat-x left bottom"">&nbsp;<img src=""images/arr.gif"" width=""11"" height=""14"" align=""absmiddle"" />&nbsp;&nbsp;产品详细介绍</td>"&VbCrLf
Response.Write "  </tr>"&VbCrLf
Response.Write "  <tr>"&VbCrLf
Response.Write "    <td height=""25"" style=""padding-left: 8px; padding-right: 8px"">"&rs("Content")&"</td>"&VbCrLf
Response.Write "  </tr>"&VbCrLf
Else
Response.Write "        <tr>"&VbCrLf
Response.Write "          <td rowspan=""2"" align=""center""><img src=""Images/NoRight.jpg""></td>"&VbCrLf
Response.Write "        </tr>"&VbCrLf
end if
rs.update
rs.close
set rs=Nothing
%>
        <tr>
          <td height="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
<%
Function Folder(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Qianbo_ProductSort where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("<center>暂无产品分类</center>")
    response.end
  end if
  i=1
  response.write("<table border='0' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from Qianbo_ProductSort where ParentID="&rs("id"))(0)
	If ISHTML = 1 Then
		AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "ProductList.Asp?SortID="&rs("id")&""
	End If
    if ChildCount=0 then
	  if i=rs.recordcount then
	    FolderType="SortFileEnd"
	  else
	    FolderType="SortFile"
	  end if
	  FolderName=rs("SortName")
	  onMouseUp=""
    else
	  if i=rs.recordcount then
	 	FolderType="SortEndFolderClose"
		ListType="SortEndListline"
		onMouseUp="EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  else
		FolderType="SortFolderClose"
		ListType="SortListline"
		onMouseUp="SortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  end if
	  FolderName=rs("SortName")
    end If
    datafrom="Qianbo_ProductSort"
    response.write("<tr>")
    response.write("<td nowrap id='b"&rs("id")&"' class='"&FolderType&"'></td><td nowrap height=23><a href="""&AutoLink&""">"&FolderName&"</a>&nbsp;")
    response.write("</td></tr>")
    if ChildCount>0 then
%>
<tr id="a<%= rs("id")%>" style="display:yes">
  <td class="<%= ListType%>" nowrap></td>
  <td ><% Folder(rs("id")) %></td>
</tr>
<%
	end if
    rs.movenext
    i=i+1
	wend
	response.write("</table>")
	rs.close
	set rs=nothing
end Function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""ProductList.asp"" class=""agray"">产品中心</a>"&VbCrLf
  if request.QueryString("ID")="" then
    WebLocation=WebLocation
  elseif not IsNumeric(request.QueryString("ID")) then
    WebLocation=WebLocation&"读取错误"
  elseif conn.execute("select * from Qianbo_Products Where ViewFlag and  ID="&request.QueryString("ID")).eof then
    WebLocation=WebLocation&"读取错误"
  else
    dim rs,sql
    set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_Products where ViewFlag and ID="&request.QueryString("ID")
    rs.open sql,conn,1,1
	WebLocation=WebLocation&SortPathTXT("Qianbo_ProductSort",rs("SortID"))
    rs.close
    set rs=nothing
  end if
end Function

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
	SortPathTXT=SortPathTXT(DataFrom,rs("ParentID"))&" - <a href="""&AutoLink&"""><span style=""color:#EC0000"">"&rs("SortName")&"</span></a>"
  end if
  rs.close
  set rs=nothing
end Function
%>