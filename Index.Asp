<!--#include file="Include/Const.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath(""&SysRootDir&"NoInstall.txt")) Then
	response.Redirect "install.asp"
End If
set fso=Nothing
Call SiteInfo
If ISHTML = 1 Then
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
<title><% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script src="Scripts/Message.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<body>
<div id="header" ><img src="images/logo.gif" width="411" height="68" style="margin:5px 0 0 10px;"  /></div>
<div id="topmenu" >
    <div id="navMenu" >
        <!-- 导航条开始 -->
        <ul>
          <%=HeadNavigation()%>
          <li class="right"></li>
        </ul>
        <!-- 导航条结束 -->
    </div>
</div>
<%
Function HeadNavigation()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_Navigation where ViewFlag order by Sequence asc"
    rs.Open sql, conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write "暂无导航"
    Else
        Do
            If ISHTML = 1 Then
                response.Write " <li><a href="""&rs("HtmlNavUrl")&""" title="""&rs("NavName")&""">"&rs("NavName")&"</a></li><li class=""line""></li>"
            Else
                response.Write " <li><a href="""&rs("NavUrl")&""" title="""&rs("NavName")&""">"&rs("NavName")&"</a></li><li class=""line""></li>"
            End If
            rs.movenext
        Loop Until rs.EOF
    End If
    rs.Close
    Set rs = Nothing
End Function
%>
<table width="1003" height="310" border="0" align="center" cellpadding="0" cellspacing="0" id="banner">
  <tr>
    <td height="310">
   	<!-- 我是首页的幻灯宣传片 -->
   	<img src="images/5073fcd656c2c.jpg" width="1000" height="323" /></td>
  </tr>
</table>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" style="background-color:#fff;">
  <tr>
    <td><table width="1003" border="0" cellspacing="0" cellpadding="0" style="margin-bottom:23px;">
        <tr>
          <td width="266" valign="top"><table width="230" border="0" cellspacing="0" cellpadding="0" style="margin:19px auto auto 12px;">
              <tr>
                <td><%=SelPlay(Video,230,210)%></td>
              </tr>
              <tr>
                <td height="21">&nbsp;</td>
              </tr>
              <tr>
                <td><img src="images/video.gif" width="230" height="27" /></td>
              </tr>
              <tr>
                <td valign="top" style="border:1px solid #DDDDDD; border-top:0px none; background:url(images/v_bg.jpg) repeat-x left bottom;"><table width="188" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="188"><img src="images/t.gif" width="1" height="10" /></td>
                    </tr>
                    <%'=videos("2,")%>
                    <tr>
                      <td height="28" align="right" valign="bottom"><a href="OtherList.asp"><img src="images/more3.gif" alt="查看更多企业视频" width="50" height="20" border="0" /></a></td>
                    </tr>
                    <tr>
                      <td height="5"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td><img src="images/links.gif" width="230" height="27" /></td>
              </tr>
              <tr>
                <td class="links" style="padding-top: 8px; padding-bottom: 8px; background:url(images/v_bg.jpg) repeat-x left bottom"><%=FriendLinks(2,2)%></td>
              </tr>
            </table></td>
          <td valign="top"><table width="486" border="0" cellspacing="0" cellpadding="0" style="margin-top:19px;">
              <tr>
                <td width="486" height="24" style="background:url(images/bg5.gif) repeat-x left bottom;"><table width="486" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="243"><img src="images/m_bt1.gif" width="120" height="24" /></td>
                      <td width="243" align="right" valign="bottom"><a href="NewsList.asp"><img src="images/more2.gif" alt="更多新闻" width="51" height="14" border="0" style="margin-bottom:7px;" /></a></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td><img src="images/t.gif" width="1" height="6" /></td>
              </tr>
              <%=News("0,")%>
            </table>
            <table width="486" border="0" cellspacing="0" cellpadding="0" style="margin-top:8px;">
              <tr>
                <td width="486" height="24" style="background:url(images/bg5.gif) repeat-x left bottom;"><table width="486" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="243"><img src="images/m_bt2.gif" width="120" height="24" /></td>
                      <td width="243" align="right" valign="bottom"><a href="ProductList.asp"><img src="images/more2.gif" alt="更多产品" width="51" height="14" border="0" style="margin-bottom:7px;" /></a></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td><img src="images/t.gif" width="1" height="6" /></td>
              </tr>
              <tr>
                <td><%=Products("0,",2,3)%></td>
              </tr>
            </table></td>
          <td width="238" valign="top"><table width="216" border="0" cellspacing="0" cellpadding="0" style="margin-top:15px; margin-left:11px;">
            <tr>
              <td height="25" style="background:url(images/bg6.gif) repeat-x left bottom;" valign="top"><span id="tag_1" class="tagdef" onclick="tag('1');">会员中心</span></td>
            </tr>
            <tr>
              <td><table cellpadding="0" cellspacing="0" id="tagCon_1">
                <tr>
                  <td><table width="216" border="0" cellspacing="0" cellpadding="0">
                    <form id="formLogin" name="formLogin" method="post" action="MemberLogin.asp">
                      <tr>
                        <td colspan="3"><img src="images/t.gif" width="1" height="10" /></td>
                      </tr>
                      <%if session("MemName")="" or session("GroupID")="" or session("MemLogin")<>"Succeed" then%>
                      <tr>
                        <td width="68" rowspan="3" align="center" valign="top"><img src="images/user_bg.gif" width="49" height="50" style="margin-top:5px;" /></td>
                        <td width="48" height="28">用户名：</td>
                        <td width="100" height="28"><input name="LoginName" type="text" id="LoginName" style="border:0px none; width:82px; background:url(images/input_bg1.jpg) repeat-x;" /></td>
                      </tr>
                      <tr>
                        <td height="28">密&nbsp;&nbsp;&nbsp;码：</td>
                        <td height="28" colspan="2"><input name="LoginPassword" type="password" id="LoginPassword" style="border:0px none; width:82px; background:url(images/input_bg1.jpg) repeat-x;" /></td>
                      </tr>
                      <tr>
                        <td height="28"></td>
                        <td height="30" colspan="2"><input type="image" name="imageField" src="images/btn_login.gif" style="border:0px none; background:none;" /></td>
                      </tr>
                      <%Else%>
                      <tr>
                        <td width="68" rowspan="3" align="center" valign="top"><img src="images/user_bg.gif" width="49" height="50" style="margin-top:5px;" /></td>
                        <td height="28" colspan="2">您好：<%=session("MemName")%></td>
                      </tr>
                      <tr>
                        <td height="28" colspan="3"><a href="MemberInfo.asp">会员中心</a> <a href="ProductBuy.asp">购物车</a> <a href="MemberLogin.asp?Action=Out">退出登录</a></td>
                      </tr>
                      <tr>
                        <td height="28" colspan="3"></td>
                      </tr>
                      <%End If%>
                    </form>
                  </table></td>
                </tr>
                <tr>
                  <td style="background:url(images/bg6.gif) repeat-x left center;"><img src="images/t.gif" width="1" height="5" /></td>
                </tr>
                <tr>
                  <td height="25" align="right" bgcolor="#9A9BA0"><img src="images/arr7.gif" width="7" height="7" align="absmiddle" />&nbsp;<a href="MemberRegister.asp">用户注册！</a>&nbsp;&nbsp;&nbsp;<img src="images/arr7.gif" width="7" height="7" align="absmiddle" /> <a href="MemberGetPass.asp">忘记密码？</a>&nbsp;&nbsp;</td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td style="background:url(images/bg6.gif) repeat-x left center;"><img src="images/t.gif" width="1" height="5" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td height="104" style="border:1px solid #999A9D;"><div style="border:1px solid #CECFD1; width:214px; height:102px; background:url(images/con_bg.jpg) repeat-x;">
                <table width="214" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="10">&nbsp;</td>
                    <td colspan="2"><img src="images/index_con_bt.gif" width="132" height="31" /></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td width="27">&nbsp;</td>
                    <td width="177" height="21">手机号码：<%=Telephone%></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td height="21">联系电话：<%=Fax%></td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td height="21" align="right" valign="bottom"><a href="About.asp"><img src="images/more3_c.gif" alt="查看详细联系方式" width="50" height="20" border="0" style="margin-right:29px;" /></a></td>
                  </tr>
                </table>
              </div></td>
            </tr>
          </table>
          	<!-- 可配置一些栏目等 -->
		  	<!--#'include file="Left.asp"-->
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="End.asp"-->
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

function News(SortPath)
  dim rs,sql,NewsName,NewFlag
  set rs = server.createobject("adodb.recordset")
  sql="select top 15 * from Qianbo_News where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
  rs.open sql,conn,1,1
  if rs.eof then
	Response.Write "  <tr>"&VbCrLf
	Response.Write "    <td align=""center"" height=""28"">暂无相关信息</td>"&VbCrLf
	Response.Write "  </tr>"&VbCrLf
  else
    do while not rs.eof
	  if now()-rs("AddTime")<=2 then
	    NewsName=StrLeft(rs("NewsName"),42)
	    NewFlag=" <img src=""Images/new.gif"" align=""absmiddle"">"
	  else
	    NewsName=StrLeft(rs("NewsName"),46)
	    NewFlag=""
	  end If
		If ISHTML = 1 Then
			AutoLink = ""&NewName&""&Separated&""&rs("ID")&"."&HTMLName&""
		Else
			AutoLink = "NewsView.asp?ID="&rs("id")&""
		End If
		Response.Write "  <tr>"&VbCrLf
		Response.Write "    <td height=""30""><table width=""486"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
		Response.Write "        <tr>"&VbCrLf
		Response.Write "          <td width=""388"" height=""30"" style=""background:url(images/bg5.gif) repeat-x left bottom;""><img src=""images/arr2.gif"" width=""3"" height=""5"" align=""absmiddle"" />&nbsp;&nbsp;<a href="""&AutoLink&""">"&NewsName&"</a>"&NewFlag&"</td>"&VbCrLf
		Response.Write "          <td width=""98"" align=""right"" style=""font-size:11px; color:#999999; background:url(images/bg5.gif) repeat-x left bottom;""><img src=""images/clock.gif"" width=""7"" height=""7"" align=""absmiddle"" />&nbsp;&nbsp;"&FormatDate(rs("Addtime"),14)&"</td>"&VbCrLf
		Response.Write "        </tr>"&VbCrLf
		Response.Write "      </table></td>"&VbCrLf
		Response.Write "  </tr>"&VbCrLf
      rs.movenext
    loop
  end if
  rs.close
  set rs=nothing
end Function

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
			Response.Write "          <td style=""border: 1px solid #ccc; width:150px; height:113px; text-align:center;""><a href="""&AutoLink&"""><img src="""&SmallPicPath&""" alt="""&rs("ProductName")&""" onload=""javascript:DrawImage(this,150,113);"" border=""0"" /></a></td>"&VbCrLf
			Response.Write "        </tr>"&VbCrLf
			Response.Write "        <tr>"&VbCrLf
			Response.Write "          <td style=""height:30px; text-align:center;""><a href="""&AutoLink&""" title="""&rs("ProductName")&""">"&ProductName&"</a> </td>"&VbCrLf
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
%>