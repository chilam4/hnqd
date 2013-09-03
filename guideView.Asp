<!--#include file="Include/Const.asp" -->
<!--#include file="Include/NoSQL.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<%
Call SiteInfo()
If ISHTML = 1 Then
    Response.expires = 0
    Response.expiresabsolute = Now() - 1
    Response.addHeader "pragma", "no-cache"
    Response.addHeader "cache-control", "private"
    Response.CacheControl = "no-cache"
End If
ID = request.QueryString("ID")
If ID <> "" Or IsNumeric(ID) Then
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_News where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 3
    If rs("SeoKeywords") <> "" Then
        SeoKeywords = rs("SeoKeywords")
    Else
        SeoKeywords = rs("NewsName")
    End If
    If rs("SeoDescription") <> "" Then
        SeoDescription = rs("SeoDescription")
    Else
        SeoDescription = rs("NewsName")
    End If
    SeoTitle = rs("NewsName")
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><% =SeoTitle %> - <% =SiteTitle %></title>
<meta name="keywords" content="<% =SeoKeywords %>" />
<meta name="description" content="<% =SeoDescription %>" />
<link href="css/public.css" rel="stylesheet" type="text/css" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div id="wrapper">
  <!--#include file="Top.asp" -->
  <div class="height7"></div>
  <div id="mainWrapper">
    <div class="content">
      <div id="innerContent">
        <div class="siteMapPath"><a href="<%=SysRootDir%>">首页</a> - <%=WebLocation()%></div>
        <div class="pageMainContent">
          <div class="tl"></div>
          <div class="tr"></div>
          <div class="tm">
            <h2 class="tt"><%=SeoTitle%></h2>
          </div>
          <div class="pack">
            <div class="ml"></div>
            <div class="mr"></div>
<%
ID = request.QueryString("ID")
If ID = "" Or Not IsNumeric(ID) Then
    response.Write "<center>暂无相关信息</center>"
ElseIf conn.Execute("select * from Qianbo_News Where ViewFlag and  ID="&ID).EOF Then
    response.Write "<center>暂无相关信息</center>"
Else
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_News where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 3
End If
If Request("page") = "" Then
    pageNum = 0
Else
    pageNum = Request("page")
End If
If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
%>
            <div class="mm">发布时间：<font color="#666666"><%=FormatDate(rs("Addtime"),13)%></font> 新闻来源：<font color="#666666"><%=rs("Source")%></font> 浏览次数：<font color="#666666">
              <script language="javascript" src="HitCount.asp?id=<%=rs("ID")%>&LX=Qianbo_News"></script>
              <script language="javascript" src="HitCount.asp?action=count&LX=Qianbo_News&id=<%=rs("ID")%>"></script>
              </font></div>
            <div class="mm">
<%
ContentStr = Split(rs("Content"), "{＄html_Paging＄}")
For i = pageNum To pageNum
%>
              <%=ProcessSitelink(ContentStr(i))%>
              <% Next %>
            </div>
            <div class="mm">本文共分
              <%For p = 0 to ubound(ContentStr)%>
              <a href="NewsView.Asp?ID=<%=request("ID")%>&page=<%=p%>"><font color="red"><%=p+1%></font></a>
              <%Next%>
              页</div>
            <%Else%>
            <div class="mm">无查看权限</div>
<%
End If
rs.update
rs.Close
Set rs = Nothing
%>
          </div>
          <div class="bl"></div>
          <div class="br"></div>
          <div class="bm"></div>
        </div>
      </div>
    </div>
    <div id="sidebar" class="sidebar">
      <div class="lframe">
        <div class="tm">
          <h2 class="tt">当前分类</h2>
        </div>
        <div class="mm">
          <ul class="nl">
            <%=WebMenu(0,0,2)%>
          </ul>
        </div>
      </div>
      <!--#include file="Left.asp" -->
    </div>
  </div>
  <!--#include file="End.asp" -->
</div>
</body>
</html>
<%
Function WebMenu(ParentID, i, level)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_NewsSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
    rs.Open sql, conn, 1, 1
    If conn.Execute("select ID from Qianbo_NewsSort Where ViewFlag and ParentID=0").EOF Then
        response.Write "<center>暂无相关信息</center>"
    End If
    Do While Not rs.EOF
        If ISHTML = 1 Then
            AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
        End If
        response.Write "<li><a href="""&AutoLink&""">"&rs("SortName")&"</a></li>"&vbCrLf
        i = i + 1
        If i<level Then Call WebMenu(rs("ID"), i, level)
        i = i -1
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
End Function

Function WebLocation()
    WebLocation = "<a href=""NewsList.Asp"" class=""agray"">公司动态</a>"&vbCrLf
    If request.QueryString("ID") = "" Then
        WebLocation = WebLocation
    ElseIf Not IsNumeric(request.QueryString("ID")) Then
        WebLocation = WebLocation&"读取错误"
    ElseIf conn.Execute("select * from Qianbo_News Where ViewFlag and  ID="&request.QueryString("ID")).EOF Then
        WebLocation = WebLocation&"读取错误"
    Else
        Dim rs, sql
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select * from Qianbo_News where ViewFlag and ID="&request.QueryString("ID")
        rs.Open sql, conn, 1, 1
        WebLocation = WebLocation&SortPathTXT("Qianbo_NewsSort", rs("SortID"))
        rs.Close
        Set rs = Nothing
    End If
End Function

Function SortPathTXT(DataFrom, ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From "&DataFrom&" where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 1
    If Not rs.EOF Then
        If ISHTML = 1 Then
            AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
        End If
        SortPathTXT = SortPathTXT(DataFrom, rs("ParentID"))&" - <a href="""&AutoLink&"""><span style=""color:#EC0000"">"&rs("SortName")&"</span></a>"
    End If
    rs.Close
    Set rs = Nothing
End Function
%>