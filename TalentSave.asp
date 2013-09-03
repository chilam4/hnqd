<!--#include file="Include/Const.asp"-->
<!--#include file="Include/ConnSiteData.asp"-->
<%
if DateDiff("s",session("time"),now())<Refresh then
   response.write "<script language='JavaScript'>alert('防刷新机制启动：请不要在 "&Refresh&" 秒内重复刷新！');" & "history.back()" & "</script>"
   response.end
else
   session("time")=now()
end if
dim rs,sql
dim JobID,TalentsName,BirthDate,Stature,Marriage,RegResidence,EduResume,JobResume
dim mMemID,mLinkman,mSex,mAddress,mZipCode,mTelephone,mMobile,mEmail,CheckCode
JobID=request.QueryString("JobID")
TalentsName=trim(request.form("TalentsName"))
BirthDate=trim(request.form("BirthDate"))
Stature=trim(request.form("Stature"))
Marriage=trim(request.form("Marriage"))
RegResidence=trim(request.form("RegResidence"))
EduResume=trim(request.form("EduResume"))
JobResume=trim(request.form("JobResume"))
mMemID=request.QueryString("MemberID")
mLinkman=trim(request.form("Linkman"))
mSex=trim(request.form("Sex"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mTelephone=trim(request.form("Telephone"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
CheckCode = Trim(request.form("CheckCode"))
dim ErrMessage,ErrMsg(14),FindErr(14),i
  ErrMsg(0)="请填写申请职位。"
  ErrMsg(1)="出生日期格式错误。"
  ErrMsg(2)="身高必须为数字。"
  ErrMsg(3)="请填写户口所在地。"
  ErrMsg(4)="请填写教育经历。"
  ErrMsg(5)="请填写工作经历。"
  ErrMsg(6)="请填写姓名。"
  ErrMsg(7)="请填写联系地址。"
  ErrMsg(8)="请正确填写邮政编码。"
  ErrMsg(9)="请正确填写联系电话。"
  ErrMsg(10)="请正确填写手机号码。"
  ErrMsg(11)="电子信箱格式错误。"
  ErrMsg(12)="验证码不能为空，请返回检查。"
  ErrMsg(13)="您在【产品订购】页面停留的时间过长，导致验证码失效。\n请返回并刷新【产品订购】页面重新订购！"
  ErrMsg(14)="您输入的验证码和系统产生的不一致，请重新输入。"
if len(TalentsName)>100 Or len(TalentsName)=0 then
  FindErr(0)=true
end if
if not IsDate(BirthDate) then
  FindErr(1)=true
end if
if not IsNumeric(Stature) Or len(Stature)=0 then
  FindErr(2)=true
end if
if len(RegResidence)>100 Or len(RegResidence)=0 then
  FindErr(3)=true
end if
if len(EduResume)=0 then
  FindErr(4)=true
end if
if len(JobResume)=0 then
  FindErr(5)=true
end if
if len(mLinkman)>50 Or len(mLinkman)=0 then
  FindErr(6)=true
end if
if len(mAddress)>100 Or len(mAddress)=0 then
  FindErr(7)=true
end if
if len(mZipCode)<>6 then
  FindErr(8)=true
end if
if len(mTelephone)>50 Or len(mTelephone)=0 then
  FindErr(9)=true
end if
if len(mMobile)>50 Or len(mMobile)<11 Or len(mMobile)=0 then
  FindErr(10)=true
end if
if not IsValidEmail(mEmail) then
  FindErr(11)=true
end If
If CheckCode = "" Then
  FindErr(12)=true
End If
If Trim(Session("CheckCode")) = "" Then
  FindErr(13)=true
End If
If CheckCode <> Session("CheckCode") Then
  FindErr(14)=true
End If
for i = 0 to UBound(FindErr)
  if FindErr(i)=true then
    ErrMessage=ErrMessage+ErrMsg(i)+"\n"
  end if
next
if not (ErrMessage="" Or isnull(ErrMessage)) then
  WriteMsg(ErrMessage)
  response.end
end if
set rs = server.createobject("adodb.recordset")
sql="select * from Qianbo_Talents"
rs.open sql,conn,1,3
rs.addnew
rs("JobID")=JobID
rs("TalentsName")=StrReplace(TalentsName)
rs("BirthDate")=BirthDate
rs("Stature")=Stature
rs("Marriage")=Marriage
rs("RegResidence")=StrReplace(RegResidence)
rs("EduResume")=StrReplace(EduResume)
rs("JobResume")=StrReplace(JobResume)
rs("MemID")=mMemID
rs("Linkman")=StrReplace(mLinkman)
rs("Sex")=mSex
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Telephone")=StrReplace(mTelephone)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("AddTime")=now()
rs.update
rs.close
set rs=nothing
conn.execute "update Qianbo_Jobs set TalentsNumber = TalentsNumber+1 where ID="&JobID
response.write "<script language='javascript'>alert('简历提交成功！');location.replace('Index.asp');</script>"
%>
