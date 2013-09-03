<%
function videos(SortPath)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 5 * from Qianbo_Others where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
  rs.open sql,conn,1,1
  if rs.eof then
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td align=""center"" height=""28"">暂无相关信息</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
  else
    do while not rs.eof
		If ISHTML = 1 Then
			AutoLink = ""&OtherName&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "OtherView.asp?ID="&rs("ID")&""
		End If
		Response.Write "  <tr>"&VbCrLf
		Response.Write "    <td height=""25"" style=""background:url(images/bg5.gif) repeat-x left bottom;""><img src=""images/arr4.gif"" width=""9"" height=""9"" align=""absmiddle"" /> <a href="""&AutoLink&""">"&StrLeft(rs("OthersName"),42)&"</a></td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
      rs.movenext
    loop
  end if
  rs.close
  set rs=nothing
end Function

function abouts()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select ID,AboutName from Qianbo_About where ViewFlag and not ChildFlag order by Sequence asc"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "<li>暂无相关信息</li>"
  else
    Do
		If ISHTML = 1 Then
			AutoLink = ""&AboutNameDiy&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "About.asp?ID="&rs("ID")&""
		End If
		response.write "<li><a href="""&AutoLink&""">"&rs("AboutName")&"</a></li>"
		rs.movenext
    loop until rs.eof
  end if
  rs.close
  set rs=nothing
end function

function Newlist(ParentID,i,level)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_NewsSort where ViewFlag order by ID asc"
  rs.open sql,conn,1,1
  if conn.execute("select ID from Qianbo_NewsSort Where ViewFlag").eof then
    response.write "<li>暂无相关信息</li>"
  end if
  do while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "NewsList.asp?SortID="&rs("ID")&""
	End If
	response.write "<li><a href="""&AutoLink&""">"&rs("SortName")&"</a></li>"
    'i=i+1
	'if i<level then call WebMenu(rs("ID"),i,level)
	'i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end function

function News(SortPath)
  dim rs,sql,NewsName,NewFlag
  set rs = server.createobject("adodb.recordset")
  sql="select top 8 * from Qianbo_News where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
  rs.open sql,conn,1,1
  if rs.eof then
	Response.Write "<li>暂无相关信息</li>"
  else
    do while not rs.eof
	  if now()-rs("AddTime")<=2 then
	    NewsName=StrLeft(rs("NewsName"),26)
	    NewFlag=" <img src=""Images/new.gif"" align=""absmiddle"">"
	  else
	    NewsName=StrLeft(rs("NewsName"),26)
	    NewFlag=""
	  end If
		If ISHTML = 1 Then
			AutoLink = ""&NewName&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "NewsView.asp?ID="&rs("id")&""
		End If
		Response.Write "<li><a href="""&AutoLink&""">"&NewsName&"</a></li>"
      rs.movenext
    loop
  end if
  rs.close
  set rs=nothing
end Function

function Joblist()
  response.write "<li><a href=""JobsList.asp"">招聘信息</a></li>"
  response.write "<li><a href=""MemberTalent.asp"">我的应聘</a></li>"
end function

function DownList(ParentID,i,level)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_DownSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
  rs.open sql,conn,1,1
  if conn.execute("select ID from Qianbo_DownSort Where ViewFlag and ParentID=0").eof then
    response.write "<center>暂无相关信息</center>"
  end if
response.write "<ul class=""nl"">"&VbCrLf
  do while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&DownSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
	Else
		AutoLink = "DownList.asp?SortID="&rs("ID")&""
	End If
	response.write "<li><a href="""&AutoLink&""">"&rs("SortName")&"</a></li>"&VbCrLf
	response.write "</ul>"&VbCrLf
    'i=i+1
	'if i<level then call WebMenu(rs("ID"),i,level)
	'i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end function

Function MenberList()
response.write "<ul class=""nl"">"&VbCrLf
If session("MemName")="" Or session("GroupID")="" Or session("MemLogin")<>"Succeed" Then
  response.write "<li><a href=""MemberRegister.asp"">注册会员</a></li>"&VbCrLf
Else
  response.write "<li><a href=""MemberInfo.asp"">修改注册资料</a></li>"&VbCrLf
  response.write "<li><a href=""MemberMessage.asp"">我的留言</a></li>"&VbCrLf
  response.write "<li><a href=""MemberOrder.asp"">我的订单</a></li>"&VbCrLf
  response.write "<li><a href=""MemberTalent.asp"">我的应聘</a></li>"&VbCrLf
  response.write "<li><a href=""MemberLogin.asp?Action=Out"">退出登录</a></li>"&VbCrLf
End If
response.write "</ul>"&VbCrLf
End Function

Function Products(SortPath,trs,tds)
  dim rs,sql,tr,td,ProductName,SmallPicPath
  set rs = server.createobject("adodb.recordset")
  sql="select top "&trs*tds&" * from Qianbo_Products where ViewFlag  order by id desc"
  'sql="select top "&trs*tds&" * from Qianbo_Products where ViewFlag and CommendFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
  'sql="select top "&trs*tds&" * from Qianbo_Products where ViewFlag and NewFlag order by id desc"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "暂无相关信息"
  else
    Response.Write "<table cellpadding=""0"" cellspacing=""0"">"&VbCrLf
	for tr=1 to trs
	    Response.Write "  <tr>"&VbCrLf
        for td=1 to tds
	      if StrLen(rs("ProductName"))<=18 then
            ProductName=rs("ProductName")
	      else
	        ProductName=StrLeft(rs("ProductName"),16)
	      end If
			If ISHTML = 1 Then
				AutoLink = ""&ProName&""&Separated&""&rs("ID")&"."&HTMLName&""
			Else
				AutoLink = "ProductView.asp?ID="&rs("id")&""
			End If
			SmallPicPath=HtmlSmallPic(rs("GroupID"),rs("SmallPic"),rs("Exclusive"))
			Response.Write "    <td><table cellpadding=""2"" cellspacing=""0"" >"&VbCrLf
			Response.Write "        <tr>"&VbCrLf
			Response.Write "          <td style=""border: 1px solid #ccc; width:130px; height:98px; text-align:center;""><a href="""&SmallPicPath&""" rel=""lightbox""><img src="""&SmallPicPath&""" alt="""&rs("ProductName")&""" onload=""javascript:DrawImage(this,130,98);"" border=""0"" /></a></td>"&VbCrLf
			Response.Write "        </tr>"&VbCrLf
			Response.Write "        <tr>"&VbCrLf
			Response.Write "          <td style=""height:20px; text-align:center;""><a href="""&AutoLink&""" title="""&rs("ProductName")&""">"&ProductName&"</a></td>"&VbCrLf
			Response.Write "        </tr><tr><td style=""height:20px; text-align:center;"">产品型号："&rs("ProductModel")&"</td></tr>"&VbCrLf
			Response.Write "      </table></td>"&VbCrLf
			Response.Write "    <td width=""8""></td>"&VbCrLf
	      rs.movenext
		  if rs.eof then exit for
		next
	    Response.Write "  </tr>"&VbCrLf
		if rs.eof then exit for
	next
    Response.Write "</table>"&VbCrLf
  end if
  rs.close
  set rs=nothing
End Function

Function FriendLinks(trs,tds)
  dim rs,sql,tr,td,ProductName,SmallPicPath
  set rs = server.createobject("adodb.recordset")
  sql="select top "&trs*tds&" * from Qianbo_FriendLink where ViewFlag order by ID desc"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write "暂无相关信息"
  else
    Response.Write "<table cellpadding=""0"" cellspacing=""0"" align=""center"">"&VbCrLf
	for tr=1 to trs
	    Response.Write "  <tr>"&VbCrLf
        for td=1 to tds
	      if StrLen(rs("LinkFace"))<=18 then
            LinkFace=rs("LinkFace")
	      else
	        LinkFace=StrLeft(rs("LinkFace"),16)
	      end if
			Response.Write "    <td><table cellpadding=""2"" cellspacing=""0"">"&VbCrLf
			Response.Write "        <tr>"&VbCrLf
			Response.Write "          <td align=""center"" height=""38"">"
			If rs("LinkType") = 0 Then
			Response.Write "<a href="""&rs("LinkUrl")&""" title="""&rs("LinkName")&""">"&LinkFace&"</a>"
			Else
			Response.Write "<a href="""&rs("LinkUrl")&""" target=""_blank""><img src="""&rs("LinkFace")&""" alt="""&rs("LinkName")&""" width=""88"" height=""31"" border=""0"" /></a>"
			End If
			Response.Write "</td>"&VbCrLf
			Response.Write "        </tr>"&VbCrLf
			Response.Write "      </table></td>"&VbCrLf
			Response.Write "    <td width=""8""></td>"&VbCrLf
	      rs.movenext
		  if rs.eof then exit for
		next
	    Response.Write "  </tr>"&VbCrLf
		if rs.eof then exit for
	next
    Response.Write "</table>"&VbCrLf
  end if
  rs.close
  set rs=nothing
End Function

Function SelPlay(strUrl,strWidth,StrHeight)
Dim Exts,isExt
If strUrl <> "" Then
   isExt = LCase(Mid(strUrl,InStrRev(strUrl, ".")+1))
Else
   isExt = ""
End If
Exts = "mp3,avi,wmv,asf,mov,rm,ra,ram,rmvb,swf"
If Instr(Exts,isExt) = 0 Then
 Response.write "<center>视频文件格式错误</center>"
Else
 Select Case isExt
  Case "mp3","avi","wmv","asf","mov"
	Response.write "<embed id=""MediaPlayer"" src="""&strUrl&""" width="""&strWidth&""" height="""&strHeight&""" loop=""0"" autostart=""false""></embed>"
  Case "mov","rm","ra","ram","rmvb"
	Response.Write "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" id=""video1"" width="""&strWidth&""" height="""&strHeight&""" viewastext>"&VbCrLf
	Response.Write "<param name=""_extentx"" value=""5503"">"&VbCrLf
	Response.Write "<param name=""_extenty"" value=""1588"">"&VbCrLf
	Response.Write "<param name=""autostart"" value=""false"">"&VbCrLf
	Response.Write "<param name=""shuffle"" value=""0"">"&VbCrLf
	Response.Write "<param name=""prefetch"" value=""0"">"&VbCrLf
	Response.Write "<param name=""nolabels"" value=""0"">"&VbCrLf
	Response.Write "<param name=""src"" value="""&strUrl&""">"&VbCrLf
	Response.Write "<param name=""controls"" value=""imagewindow,statusbar,controlpanel"">"&VbCrLf
	Response.Write "<param name=""console"" value=""raplayer"">"&VbCrLf
	Response.Write "<param name=""loop"" value=""0"">"&VbCrLf
	Response.Write "<param name=""numloop"" value=""0"">"&VbCrLf
	Response.Write "<param name=""center"" value=""0"">"&VbCrLf
	Response.Write "<param name=""maintainaspect"" value=""0"">"&VbCrLf
	Response.Write "<param name=""backgroundcolor"" value=""#000000"">"&VbCrLf
	Response.Write "</object>"&VbCrLf
  Case "swf"
	Response.Write "<object codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width="""&strWidth&""" height="""&strHeight&""" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"&VbCrLf
	Response.Write "<param name=""_cx"" value=""19050"">"&VbCrLf
	Response.Write "<param name=""_cy"" value=""3863"">"&VbCrLf
	Response.Write "<param name=""FlashVars"" value="""">"&VbCrLf
	Response.Write "<param name=""Movie"" value="""&strUrl&""">"&VbCrLf
	Response.Write "<param name=""Src"" value="""&strUrl&""">"&VbCrLf
	Response.Write "<param name=""Play"" value=""-1"">"&VbCrLf
	Response.Write "<param name=""Loop"" value=""-1"">"&VbCrLf
	Response.Write "<param name=""Quality"" value=""High"">"&VbCrLf
	Response.Write "<param name=""SAlign"" value="""">"&VbCrLf
	Response.Write "<param name=""Menu"" value=""0"">"&VbCrLf
	Response.Write "<param name=""Base"" value="""">"&VbCrLf
	Response.Write "<param name=""AllowScriptAccess"" value="""">"&VbCrLf
	Response.Write "<param name=""Scale"" value=""ShowAll"">"&VbCrLf
	Response.Write "<param name=""DeviceFont"" value=""0"">"&VbCrLf
	Response.Write "<param name=""EmbedMovie"" value=""0"">"&VbCrLf
	Response.Write "<param name=""BGColor"" value="""">"&VbCrLf
	Response.Write "<param name=""SWRemote"" value="""">"&VbCrLf
	Response.Write "<param name=""MovieData"" value="""">"&VbCrLf
	Response.Write "<param name=""SeamlessTabbing"" value=""1"">"&VbCrLf
	Response.Write "<param name=""Profile"" value=""0"">"&VbCrLf
	Response.Write "<param name=""ProfileAddress"" value="""">"&VbCrLf
	Response.Write "<param name=""ProfilePort"" value=""0"">"&VbCrLf
	Response.Write "<embed src="""&strUrl&""" width="""&strWidth&""" height="""&strHeight&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" menu=""false""></embed>"&VbCrLf
	Response.Write "</object>"&VbCrLf
 End Select
End If
End Function

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
%>