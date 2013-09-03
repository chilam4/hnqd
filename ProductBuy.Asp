<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%call SiteInfo%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>订购产品 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script language="javascript" src="Scripts/Html.js"></script>
</head>
<!--#include file="Top.asp" -->
<%
dim Products
dim mMemID,mRealName,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail
if session("MemName")<>"" and session("MemLogin")="Succeed" then
  call MemInfo()
else
  mSex="先生"
  mMemID=0
end if
%>
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
          <td><% call ProductList() %></td>
        </tr>
        <tr>
          <td><table width="95%" border="0" align="center" cellpadding="3" cellspacing="5">
              <form action="ProductBuySave.asp?MemberID=<%=mMemID%>" method="post" name="formBuy" id="formBuy">
                <tr>
                  <td width="200" align="right">标题：</td>
                  <td><input name="OrderName" type="text" id="OrderName" value="订购产品" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">相关说明：</td>
                  <td><input type="hidden" name="Products" value="<%=Products%>" />
                    <textarea name="Remark" cols="80" rows="8" id="Remark"></textarea></td>
                </tr>
                <tr>
                  <td align="right">联系人：</td>
                  <td><input name="RealName" type="text" id="RealName" value="<%=mRealName%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">性别：</td>
                  <td><input name="Sex" type="radio" value="先生" class="inputnoborder" <%if mSex="先生" then response.write ("checked")%> />
                    先生
                    <input type="radio" name="Sex" value="女士" class="inputnoborder" <%if mSex="女士" then response.write ("checked")%> />
                    女士 <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">单位名称/：</td>
                  <td><input name="Company" type="text" id="Company" value="<%=mCompany%>" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">详细地址：</td>
                  <td><input name="Address" type="text" id="Address" value="<%=mAddress%>" size="40" maxlength="100" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">邮政编码：</td>
                  <td><input name="ZipCode" type="text" id="ZipCode" value="<%=mZipCode%>" size="20" maxlength="20" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">联系电话：</td>
                  <td><input name="Telephone" type="text" id="Telephone" value="<%=mTelephone%>" size="20" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">传真号码：</td>
                  <td><input name="Fax" type="text" id="Fax" value="<%=mFax%>" size="20" maxlength="50" /></td>
                </tr>
                <tr>
                  <td align="right">手机号码：</td>
                  <td><input name="Mobile" type="text" id="Mobile" value="<%=mMobile%>" size="20" maxlength="50" /></td>
                </tr>
                <tr>
                  <td align="right">电子信箱：</td>
                  <td><input name="Email" type="text" id="Email" value="<%=mEmail%>" size="30" maxlength="50" />
                    <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">验证码：</td>
                  <td><input name="CheckCode" type="text" size="6" maxlength="6" />
                    <a href="javascript:refreshimg()" title="看不清楚，换个图片。"><img src="Include/CheckCode/CheckCode.asp" name="checkcode" align="absmiddle" id="checkcode" style="border: 1px solid #ffffff" /></a> <font color="red">*</font></td>
                </tr>
                <tr>
                  <td align="right">&nbsp;</td>
                  <td><input name="Submit" type="submit" value="立即订购以上产品" /></td>
                </tr>
              </form>
            </table></td>
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
    i=i+1
	if i<level then call WebMenu(rs("ID"),i,level)
	i=i-1
	rs.movenext
  loop
  rs.close
  set rs=nothing
end function

function WebLocation()
  WebLocation="&nbsp;当前位置：<a href=""index.asp"" class=""agray"">首页</a> - <a href=""ProductList.asp"" class=""agray"">公司产品</a> - 订单"&VbCrLf
end Function

function NoList()
  if request("UpdateOrder")="更新选择" then
    Session("NoList")=""
  end if
  dim ProductNo,NoArray,i
  if trim(request("ProductNo"))<>"" then ProductNo=trim(request("ProductNo"))&","
  if session("NoList")="" and trim(request("ProductNo"))="" then
    NoList=""
    exit function
  end if
  if instr(ProductNo,",")>0 then
  	NoArray=split(ProductNo, ",")
  	for i = 0 to ubound(NoArray)
      if not instr(session("NoList"),NoArray(i)&",")>0 then session("NoList")=session("NoList")&NoArray(i)&","
  	next
  end if
  for i = 0 to ubound(split(session("NoList"), ","))
    NoList=NoList&"'"&trim(split(session("NoList"),",")(i))&"',"
  next
  NoList=left(NoList,len(NoList)-4)
end Function

function ProductList()
	Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&VbCrLf
	Response.Write "  <form action=""ProductBuy.asp"" method=""POST"" name=""check"">"&VbCrLf
	Response.Write "    <tr>"&VbCrLf
	Response.Write "      <td height=""28"" colspan=""2"" style=""background:url(Images/bg2.gif) repeat-x left bottom;"">您选购的产品如下：</td>"&VbCrLf
	Response.Write "    </tr>"&VbCrLf
if Nolist()="" then
	Response.Write "    <tr>"&VbCrLf
	Response.Write "      <td height=""28"" colspan=""2"" align=""center"">未选择任何产品</td>"&VbCrLf
	Response.Write "    </tr>"&VbCrLf
else
	dim rs,sql
	set rs = server.createobject("adodb.recordset")
	sql="select * from Qianbo_Products where ProductNo in ("&NoList()&") order by id"
	rs.open sql,conn,1,1
	while not rs.eof
	If ISHTML = 1 Then
		AutoLink = ""&ProName&""&Separated&""&rs("ID")&"."&HTMLName&""
	Else
		AutoLink = "ProductView.asp?ID="&rs("ID")&""
	End If
	Response.Write "    <tr>"&VbCrLf
	Response.Write "      <td width=""3%"" height=""28"" style=""background:url(Images/bg2.gif) repeat-x left bottom;""><input type=""CheckBox"" name=""ProductNo"" value="""&rs("ProductNo")&""" class=""inputnoborder"" Checked></td>"&VbCrLf
	Response.Write "      <td width=""97%"" style=""background:url(Images/bg2.gif) repeat-x left bottom;""><a href="""&AutoLink&""" title="""&rs("ProductName")&""">"&rs("ProductName")&"&nbsp;["&rs("ProductNo")&"]</a></td>"&VbCrLf
	Response.Write "    </tr>"&VbCrLf
	Products=Products&rs("ProductName")&" 产品编号："&rs("ProductNo")&"<br />"
	rs.movenext
    wend
	Response.Write "    <tr>"&VbCrLf
	Response.Write "      <td colspan=""2"" height=""35""><input type=""submit"" name=""UpdateOrder"" value=""更新选择""></td>"&VbCrLf
	Response.Write "    </tr>"&VbCrLf
    rs.close
    set rs=nothing
end if
	Response.Write "  </form>"&VbCrLf
	Response.Write "</table>"&VbCrLf
end function

sub MemInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_Members where MemName='"&session("MemName")&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "<center>暂无相关信息</center>"
  else
    mMemID=rs("ID")
	if not rs("RealName")="" then
	  mRealName=rs("RealName")
	else
	  mRealName=rs("MemName")
	end if
	mSex=rs("Sex")
	mCompany=rs("Company")
	mAddress=rs("Address")
	mZipCode=rs("ZipCode")
	mTelephone=rs("Telephone")
	mFax=rs("Fax")
	mMobile=rs("Mobile")
	mEmail=rs("Email")
  end if
  rs.close
  set rs=nothing
end sub
%>
<script language="javascript">
<!--
function refreshimg(){document.all.checkcode.src="Include/CheckCode/CheckCode.asp";}
//-->
</script>