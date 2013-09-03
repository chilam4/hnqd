<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<body>
<div id="header"><img src="images/logo.gif" width="411" height="68" style="margin:5px 0 0 10px;"  /></div>
<div id="topmenu" >
    <div id="navMenu">
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