<%
Function ProcessSitelink(ContentHTML)
    If GetCache("SitelinkState") = "No" Then ProcessSitelink = ContentHTML: Exit Function
    If Not ChkCache("Sitelink") Then
        Dim Rslink
        Set Rslink = conn.Execute("Select [Text],[Link],[Replace],[Target],[Description] From [Qianbo_Link] Where [State] = 1 Order By [Order] Desc")
        If Not Rslink.EOF Then
            Call SetCache("Sitelink", Rslink.Getrows())
            Call SetCache("SitelinkState", "Yes")
        Else
            Call SetCache("SitelinkState", "No")
            ProcessSitelink = ContentHTML
            Exit Function
        End If
        Rslink.Close: Set Rslink = Nothing
    End If
    Dim RegEx, Matches, Match
    Set RegEx = New RegExp
    RegEx.Ignorecase = True
    RegEx.Global = True
    Dim Dat, i, j, Url, UrlTitle
    Dat = GetCache("Sitelink")
    For i = 0 To UBound(Dat, 2)
        j = 0
        If InStr(ContentHTML, Dat(0, i)) > 0 Then
            RegEx.Pattern = "(>[^><]*)" & Dat(0, i) & "([^><]*<)(?!/a)"
            Set Matches = RegEx.Execute(">" & ContentHTML & "<")
            For Each Match In Matches
                UrlTitle = Dat(4, i)
                If InStr(UrlTitle, "|") > 0 Then Randomize: UrlTitle = Split(UrlTitle, "|")(Round(UBound(Split(UrlTitle, "|")) * Rnd))
                If Dat(3, i) = 1 Then Url = "<a href='" & Dat(1, i) & "' title='" & UrlTitle & "' target='_blank'>" & Dat(0, i) & "</a>" Else Url = "<a href='" & Dat(1, i) & "' title='" & UrlTitle & "'>" & Dat(0, i) & "</a>"
                Url = Replace(Url, "$", "&#36;")
                ContentHTML = Replace(ContentHTML, Match.Value, Match.SubMatches(0) & Url & Match.SubMatches(1))
                j = j + 1: If Dat(2, i) > 0 And j >= Dat(2, i) Then Exit For
            Next
        End If
    Next
    ProcessSitelink = ContentHTML
End Function
function setcache(cachename,cachevalue)
	dim cachedata
	cachename = lcase(filterstr(cachename))
	cachedata = application(Cacheflag & cachename)
	if isarray(cachedata) then
		cachedata(0) = Cachevalue
		cachedata(1) = now()
	else
		Redim cachedata(2)
		cachedata(0) = Cachevalue
		cachedata(1) = now()
	end if
	application.lock
	application(Cacheflag & cachename) = cachedata
	application.unlock
end function
function getcache(cachename)
	dim cachedata
	cachename = lcase(filterstr(cachename))
	cachedata = application(Cacheflag & cachename)
	if isarray(cachedata) then getcache = cachedata(0) else getcache = ""
end function
function chkcache(cachename)
	dim cachedata
	chkcache = false
	cachename = lcase(filterstr(cachename))
	cachedata = application(Cacheflag & cachename)
	if not isarray(cachedata) then exit function
	if not IsDate(cachedata(1)) then exit function
	if DateDiff("s", CDate(cachedata(1)), now()) < 60 * Cachetime then chkcache = true
end Function
function filterstr(str)
	filterstr = lcase(str) : filterstr = replace(filterstr, " ", "") : filterstr = replace(filterstr, "'", "") : filterstr = replace(filterstr, """", "") : filterstr = replace(filterstr, "=", "") : filterstr = replace(filterstr, "*", "")
end function

public SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,Video,IcpNumber,MesViewFlag,SiteDetail,SiteLogo
sub SiteInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from Qianbo_Site"
  rs.open sql,conn,1,1
  SiteTitle=rs("SiteTitle")
  Keywords=rs("Keywords")
  Descriptions=rs("Descriptions")
  SiteUrl=rs("SiteUrl")
  ComName=rs("ComName")
  Address=rs("Address")
  ZipCode=rs("ZipCode")
  Telephone=rs("Telephone")
  Fax=rs("Fax")
  Email=rs("Email")
  Video=rs("Video")
  IcpNumber=rs("IcpNumber")
  SiteLogo=rs("SiteLogo")
  SiteDetail=rs("SiteDetail")
  MesViewFlag=rs("MesViewFlag")
  rs.close
  set rs=nothing
end Sub

Function QianboEntCoder(password, QianboEntCode)
 MIN_Morfi = 32
 MAX_Morfi = 126
 NUM_Morfi = MAX_Morfi - MIN_Morfi + 1
    offset = password
    Rnd -1
    Randomize offset
	QianboEntCode = Replace(QianboEntCode, Chr(-23646), Chr(34))
    str_len = Len(QianboEntCode)
    For i = 1 To str_len
        ch = Asc(Mid(QianboEntCode, i, 1))
        If ch >= MIN_Morfi And ch <= MAX_Morfi Then
            ch = ch - MIN_Morfi
            offset = Int((NUM_Morfi + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_Morfi)
            If ch < 0 Then ch = ch + NUM_Morfi
            ch = ch + MIN_Morfi
            to_text = to_text & Chr(ch)
             QianboEntCoder = Replace(to_text, "|^|", vbCrLf)
        End If
    Next
End Function
Function Checkfile(filePath)
Dim fso,path
Set fso = Server.CreateObject("scripting.filesystemobject")
If InStr(filePath, ":") <> 0 Then path = filePath Else path = Server.MapPath(filePath)
If fso.fileexists(path) Then Checkfile = True Else Checkfile = False
Set fso = Nothing
End Function
Function ReadText(FileName)
Dim adf,path
Set adf = Server.CreateObject("Adodb.Stream")
If InStr(FileName, ":\") <> 0 Then path = FileName Else path = Server.MapPath(FileName)
With adf
.Type = 2
.LineSeparator = 13
.Open
.LoadFromFile (path)
.Charset = "GB2312"
.Position = 2
ReadText = .ReadText
.Cancel
.Close
End With
Set adf = Nothing
End Function
Function EntMDBr(Code)
code=replace(code,"＂", Chr(34))
dim Morfi1, Morfi2, Morfi3,i
          Morfi1 = Len(code)
          For i = 0 To Morfi1 - 1
                  Morfi2 = Asc(Right(code, Morfi1 - i)) Xor 20
                  Morfi3 =Morfi3 + Chr(Int(Morfi2))
          Next
            EntMDBr = replace(Morfi3, "§",vbcrlf)
End Function
function StrLen(Str)
  if Str="" or isnull(Str) then
    StrLen=0
    exit function
  else
    dim regex
    set regex=new regexp
    regEx.Pattern ="[^\x00-\xff]"
    regex.Global =true
    Str=regEx.replace(Str,"^^")
    set regex=nothing
    StrLen=len(Str)
  end if
end function

function StrLeft(Str,StrLen)
  dim L,T,I,C
  if Str="" then
    StrLeft=""
    exit function
  end if
  Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
  L=Len(Str)
  T=0
  for i=1 to L
    C=Abs(AscW(Mid(Str,i,1)))
    if C>255 then
      T=T+2
    else
      T=T+1
    end if
    if T>=StrLen then
      StrLeft=Left(Str,i) & ""
      exit for
    else
      StrLeft=Str
    end if
  next
  StrLeft=Replace(Replace(Replace(replace(StrLeft," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

function StrReplace(Str)
  if Str="" or isnull(Str) then
    StrReplace=""
    exit function
  else
    StrReplace=replace(str," ","&nbsp;")
    StrReplace=replace(StrReplace,chr(13),"&lt;br&gt;")
    StrReplace=replace(StrReplace,"<","&lt;")
    StrReplace=replace(StrReplace,">","&gt;")
  end if
end function
function ReStrReplace(Str)
  if Str="" or isnull(Str) then
    ReStrReplace=""
    exit function
  else
    ReStrReplace=replace(Str,"&nbsp;"," ")
    ReStrReplace=replace(ReStrReplace,"<br />",chr(13))
    ReStrReplace=replace(ReStrReplace,"&lt;br&gt;",chr(13))
    ReStrReplace=replace(ReStrReplace,"&lt;","<")
    ReStrReplace=replace(ReStrReplace,"&gt;",">")
  end if
end function
function HtmlStrReplace(Str)
  if Str="" or isnull(Str) then
    HtmlStrReplace=""
    exit function
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br />")
  end if
end function
function ViewNoRight(GroupID,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  ViewNoRight=true
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then
	    ViewNoRight=false
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	    ViewNoRight=false
      end if
  end select
end function
Function GetUrl()
  GetUrl="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
  If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL=GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function

function HtmlSmallPic(GroupID,PicPath,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  HtmlSmallPic=PicPath
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
    case "="
      if not session("GroupLevel") = GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
  end select
  if HtmlSmallPic="" or isnull(HtmlSmallPic) then HtmlSmallPic="../Images/NoPicture.jpg"
end function
function IsValidMemName(memname)
  dim i, c
  IsValidMemName = true
  if not (3<=len(memname) and len(memname)<=16) then
    IsValidMemName = false
    exit function
  end if
  for i = 1 to Len(memname)
    c = Mid(memname, i, 1)
    if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 and not IsNumeric(c) then
      IsValidMemName = false
      exit function
    end if
  next
end function
function IsValidEmail(email)
  dim names, name, i, c
  IsValidEmail = true
  names = Split(email, "@")
  if UBound(names) <> 1 then
    IsValidEmail = false
    exit function
  end if
  for each name in names
	if Len(name) <= 0 then
	  IsValidEmail = false
      exit function
    end if
    for i = 1 to Len(name)
      c = Mid(name, i, 1)
      if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 and not IsNumeric(c) then
        IsValidEmail = false
        exit function
      end if
	next
	if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
    end if
  next
  if InStr(names(1), ".") <= 0 then
    IsValidEmail = false
    exit function
  end if
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then
    IsValidEmail = false
    exit function
  end if
  if InStr(email, "..") > 0 then
    IsValidEmail = false
  end if
end function
Function FormatDate(DateAndTime, Format)
  On Error Resume Next
  Dim yy,y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(Format) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  yy = CStr(Year(DateAndTime))
  y = Mid(CStr(Year(DateAndTime)),3)
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case Format
  Case "1"
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
    strDateTime = yy & m & d & h & mi & s
  Case "3"
    strDateTime = yy & m & d & h & mi
  Case "4"
    strDateTime = yy & "年" & m & "月" & d & "日"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "月" & d & "日"
  Case "8"
    strDateTime = y & "年" & m & "月"
  Case "9"
    strDateTime = y & "-" & m
  Case "10"
    strDateTime = y & "/" & m
  Case "11"
    strDateTime = y & "-" & m & "-" & d
  Case "12"
    strDateTime = y & "/" & m & "/" & d
  Case "13"
    strDateTime = yy & "." & m & "." & d
  Case "14"
    strDateTime = yy & "-" & m & "-" & d
  Case Else
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function
function WriteMsg(Message)
  response.write "<script language='JavaScript'>alert('"&Message&"');" & "history.back()" & "</script>"
end function

Function CheckStr(Strer,Num)
	Dim Shield,w
	If Strer = "" Or IsNull(Strer) Then Exit Function
	Select Case Num
	  Case 1
		If IsNumeric(Strer) = 0 Then
		  Response.Write "操作错误"
		  Response.End
		End If
		Strer = Int(Strer)
	End Select
	CheckStr = Strer
End Function
%>