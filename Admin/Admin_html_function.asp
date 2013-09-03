<%
Sub HtmlProSort
totalrec=Conn.Execute("Select count(*) from Qianbo_Products Where ViewFlag")(0)
totalpage=int(totalrec/ProInfo)
If (totalpage * ProInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&ProSortName&""&Separated&"1."&HTMLName&"","ProductList.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&ProSortName&""&Separated&""&i&"."&HTMLName&"","ProductList.asp","Page=",i,"","")
next
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Qianbo_ProductSort order by ID desc"
rs.open sql,conn,1,1
If rs.eof Then
	Class_Num=0
Else
	Class_Num=1
do while not rs.eof
ID=rs("ID")
SortPath=conn.execute("select * from Qianbo_ProductSort Where ViewFlag And ID="&ID)("SortPath")
totalrec=Conn.Execute("Select count(*) from Qianbo_Products where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
totalpage=int(totalrec/ProInfo)
If (totalpage * ProInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&ProSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"","ProductList.asp","SortID=",ID,"Page=",1)
else
for i=1 to totalpage
call htmll("","",""&ProSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"","ProductList.asp","SortID=",ID,"Page=",i)
next
end If
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
End Sub

Sub HtmlOtherSort
totalrec=Conn.Execute("Select count(*) from Qianbo_Others Where ViewFlag")(0)
totalpage=int(totalrec/OtherInfo)
If (totalpage * OtherInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&OtherSortName&""&Separated&"1."&HTMLName&"","OtherList.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&OtherSortName&""&Separated&""&i&"."&HTMLName&"","OtherList.asp","Page=",i,"","")
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Qianbo_OthersSort order by ID desc"
rs.open sql,conn,1,1
If rs.eof Then
	Class_Num=0
Else
	Class_Num=1
do while not rs.eof
ID=rs("ID")
SortPath=conn.execute("select * from Qianbo_OthersSort Where ViewFlag And ID="&ID)("SortPath")
totalrec=Conn.Execute("Select count(*) from Qianbo_Others where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
totalpage=int(totalrec/OtherInfo)
If (totalpage * OtherInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&OtherSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"","OtherList.asp","SortID=",ID,"Page=",1)
else
for i=1 to totalpage
call htmll("","",""&OtherSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"","OtherList.asp","SortID=",ID,"Page=",i)
next
end If
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
End Sub

Sub HtmlNewSort
totalrec=Conn.Execute("Select count(*) from Qianbo_News Where ViewFlag")(0)
totalpage=int(totalrec/NewInfo)
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&NewSortName&""&Separated&"1."&HTMLName&"","NewsList.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&NewSortName&""&Separated&""&i&"."&HTMLName&"","NewsList.asp","Page=",i,"","")
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Qianbo_NewsSort order by ID desc"
rs.open sql,conn,1,1
If rs.eof Then
	Class_Num=0
Else
	Class_Num=1
do while not rs.eof
ID=rs("ID")
SortPath=conn.execute("select * from Qianbo_NewsSort Where ViewFlag And ID="&ID)("SortPath")
totalrec=Conn.Execute("Select count(*) from Qianbo_News where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)

totalpage=int(totalrec/NewInfo)
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&NewSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"","NewsList.asp","SortID=",ID,"Page=",1)
else
for i=1 to totalpage
call htmll("","",""&NewSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"","NewsList.asp","SortID=",ID,"Page=",i)
next
End If
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
End Sub

Sub HtmlJobSort
totalrec=Conn.Execute("Select count(*) from Qianbo_Jobs Where ViewFlag")(0)
totalpage=int(totalrec/JobInfo)
If (totalpage * JobInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&JobSortName&""&Separated&"1."&HTMLName&"","JobsList.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&JobSortName&""&Separated&""&i&"."&HTMLName&"","JobsList.asp","Page=",i,"","")
Response.Write "<script>bar_img.width="&Fix((i/totalpage)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & formatnumber(i/totalpage*100) & """;</script>"
Response.Flush
next
end if
End Sub

Sub HtmlDownSort
totalrec=Conn.Execute("Select count(*) from Qianbo_Download Where ViewFlag")(0)
totalpage=int(totalrec/DownInfo)
If (totalpage * DownInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&DownSortName&""&Separated&"1."&HTMLName&"","DownList.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&DownSortName&""&Separated&""&i&"."&HTMLName&"","DownList.asp","Page=",i,"","")
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Qianbo_DownSort order by ID desc"
rs.open sql,conn,1,1
If rs.eof Then
	Class_Num=0
Else
	Class_Num=1
do while not rs.eof
ID=rs("ID")
SortPath=conn.execute("select * from Qianbo_DownSort Where ViewFlag And ID="&ID)("SortPath")
totalrec=Conn.Execute("Select count(*) from Qianbo_Download where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
totalpage=int(totalrec/DownInfo)
If (totalpage * DownInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&DownSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"","DownList.asp","SortID=",ID,"Page=",1)
else
for i=1 to totalpage
call htmll("","",""&DownSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"","DownList.asp","SortID=",ID,"Page=",i)
next
end If
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
End Sub

Sub HtmlPro
totalrec=Conn.Execute("select count(*) from Qianbo_Products where Stock>0")(0)
sql="Select * from Qianbo_Products where Stock>0 order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&ProName&""&Separated&""&ID&"."&HTMLName&"","ProductView.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlOther
totalrec=Conn.Execute("select count(*) from Qianbo_Others")(0)
sql="Select * from Qianbo_Others order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&OtherName&""&Separated&""&ID&"."&HTMLName&"","OtherView.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlNews
totalrec=Conn.Execute("select count(*) from Qianbo_News")(0)
sql="Select * from Qianbo_News order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&NewName&""&Separated&""&ID&"."&HTMLName&"","NewsView.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlJob
totalrec=Conn.Execute("select count(*) from Qianbo_Jobs")(0)
sql="Select * from Qianbo_Jobs order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&JobNameDiy&""&Separated&""&ID&"."&HTMLName&"","JobsView.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlInfo
totalrec=Conn.Execute("select count(*) from Qianbo_About")(0)
sql="Select * from Qianbo_About order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&AboutNameDiy&""&Separated&""&ID&"."&HTMLName&"","About.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlDown
totalrec=Conn.Execute("select count(*) from Qianbo_Download")(0)
sql="Select * from Qianbo_Download order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
call htmll("","",""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"","DownView.asp","ID=",ID,"","")
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
End Sub

Sub HtmlIndex
totalrec=Conn.Execute("Select count(*) from Qianbo_Products Where ViewFlag")(0)
totalpage=int(totalrec/ProInfo)
If (totalpage * ProInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("","",""&AdvisoryNameDiy&""&Separated&"1."&HTMLName&"","ProductAdvisory.asp","Page=",1,"","")
else
for i=1 to totalpage
call htmll("","",""&AdvisoryNameDiy&""&Separated&""&i&"."&HTMLName&"","ProductAdvisory.asp","Page=",i,"","")
Response.Write "<script>bar_img.width="&Fix((i/totalpage)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & formatnumber(i/totalpage*100) & """;</script>"
Response.Flush
next
end if
conn.close
set conn=Nothing
call htmll("","","Index."&HTMLname&"","Index.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((1/8)*300)&";bar_txt1.innerHTML=""成功生成首页。完成比例" & formatnumber(1/8*100) & """;</script>"
Response.Flush
call htmll("","",""&AboutNameDiy&"."&HTMLname&"","Company.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((2/8)*300)&";bar_txt1.innerHTML=""成功生成“关于我们”静态页面。完成比例：" & formatnumber(2/8*100) & """;</script>"
Response.Flush
call htmll("","",""&NewSortName&"."&HTMLname&"","NewsList.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((3/8)*300)&";bar_txt1.innerHTML=""成功生成“新闻列表”静态页面。完成比例：" & formatnumber(3/8*100) & """;</script>"
Response.Flush
call htmll("","",""&ProSortName&"."&HTMLname&"","ProductList.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((4/8)*300)&";bar_txt1.innerHTML=""成功生成“产品列表”静态页面。完成比例：" & formatnumber(4/8*100) & """;</script>"
Response.Flush
call htmll("","",""&JobSortName&"."&HTMLname&"","JobsList.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((5/8)*300)&";bar_txt1.innerHTML=""成功生成“人才列表”静态页面。完成比例：" & formatnumber(5/8*100) & """;</script>"
Response.Flush
call htmll("","",""&DownSortName&"."&HTMLname&"","DownList.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((6/8)*300)&";bar_txt1.innerHTML=""成功生成“下载列表”静态页面。完成比例：" & formatnumber(6/8*100) & """;</script>"
Response.Flush
call htmll("","",""&OtherSortName&"."&HTMLname&"","OtherList.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((7/8)*300)&";bar_txt1.innerHTML=""成功生成“其他信息列表”静态页面。完成比例：" & formatnumber(7/8*100) & """;</script>"
Response.Flush
End Sub
%>