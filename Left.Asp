<table width="216" border="0" cellspacing="0" cellpadding="0" style="margin-left:11px; margin-top:17px;">
  <tr>
    <td><img src="images/right_bt3.gif" width="216" height="27" /></td>
  </tr>
  <tr>
    <td height="60" style="border:1px solid #DDDDDD; border-top:0px none; background:url(images/v_bg.jpg) repeat-x left bottom;"><form id="Search" name="Search" method="get" action="Search.asp">
        <table width="188" border="0" align="center" cellpadding="0" cellspacing="0" style="background:url(images/bg_sear.gif) no-repeat 2px 24px;">
          <tr>
            <td width="188"><img src="images/t.gif" width="1" height="9" /></td>
          </tr>
          <tr>
            <td height="28"><input name="Range" type="radio" value="Product" class="inputnoborder" checked="checked" />产品 <input type="radio" name="Range" value="News" class="inputnoborder" />新闻 <input type="radio" name="Range" value="Down" class="inputnoborder" />下载</td>
          </tr>
          <tr>
            <td height="28">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="Keyword" type="text" id="Keyword" style="width:102px;" /></td>
          </tr>
          <tr>
            <td height="28">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="image" name="imageField3" src="images/btn_sear.gif" class="inputnoborder" /></td>
          </tr>
          <tr>
            <td><img src="images/t.gif" width="1" height="9" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
  </tr>
  <tr>
    <td height="50"><img src="images/office.jpg" alt="广告位" border="0" /></a></td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
  </tr>
  <tr>
    <td width="216"><img src="images/right_bt2.gif" width="216" height="27" /></td>
  </tr>
  <tr>
    <td><div id="quick"><%=LeftNavigation%></div></td>
  </tr>
</table>
<%
function LeftNavigation()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Navigation where ViewFlag and NavName <> '首页' order by Sequence asc"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "暂无导航"
  else
	do
		If ISHTML = 1 Then
		response.write " <a href="""&rs("HtmlNavUrl")&""" title="""&rs("NavName")&""">"&rs("NavName")&"</a>"
		Else
		response.write " <a href="""&rs("NavUrl")&""" title="""&rs("NavName")&""">"&rs("NavName")&"</a>"
		End If
	  rs.movenext
	loop until rs.eof
  end if
  rs.close
  set rs=nothing
end Function
%>