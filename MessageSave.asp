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
dim MesName,Content,SecretFlag,mMemID,mLinkman,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,CheckCode
MesName=trim(request.form("MesName"))
Content=trim(request.form("Content"))
if trim(request.form("SecretFlag"))="1" then
  SecretFlag=1
else
  SecretFlag=0
end if
mMemID=request.QueryString("MemberID")
mLinkman=trim(request.form("Linkman"))
mSex=trim(request.form("Sex"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mTelephone=trim(request.form("Telephone"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
CheckCode = Trim(request.form("CheckCode"))
dim ErrMessage,ErrMsg(9),FindErr(9),i
  ErrMsg(0)="请填写留言主题。"
  ErrMsg(1)="请正确填写留言内容。"
  ErrMsg(2)="请填写您的称呼。"
  ErrMsg(3)="请正确填写单位名称、联系地址。"
  ErrMsg(4)="请正确填写邮政编码。"
  ErrMsg(5)="请正确填写联系电话、传真号码、手机号码。"
  ErrMsg(6)="电子信箱格式错误。"
  ErrMsg(7)="验证码不能为空，请返回检查。"
  ErrMsg(8)="您在【会员注册】页面停留的时间过长，导致验证码失效。\n请返回并刷新【会员注册】页面重新注册！"
  ErrMsg(9)="您输入的验证码和系统产生的不一致，请重新输入。"
if len(MesName)>100 Or len(MesName)=0 then
  FindErr(0)=true
end if
if len(Content)<10 then
  FindErr(1)=true
end if
if len(mLinkman)>50 Or len(mLinkman)=0 then
  FindErr(2)=true
end if
if len(mCompany)>100 Or len(mAddress)>100 Or len(mCompany)=0 Or len(mAddress)=0 then
  FindErr(3)=true
end if
if len(mZipCode)<>6 then
  FindErr(4)=true
end if
if len(mTelephone)>50 Or len(mFax)>50 Or len(mMobile)>50 Or len(mTelephone)=0 Or len(mFax)=0 Or len(mMobile)=0 then
  FindErr(5)=true
end if
if not IsValidEmail(mEmail) then
  FindErr(6)=true
end If
If CheckCode = "" Then
	FindErr(7)=true
End If
If Trim(Session("CheckCode")) = "" Then
	FindErr(8)=true
End If
If CheckCode <> Session("CheckCode") Then
	FindErr(9)=true
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
sql="select * from Qianbo_Message"
rs.open sql,conn,1,3
rs.addnew
rs("MesName")=StrReplace(MesName)
rs("Content")=StrReplace(Content)
rs("MemID")=mMemID
rs("Linkman")=StrReplace(mLinkman)
rs("Sex")=mSex
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Telephone")=StrReplace(mTelephone)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("SecretFlag")=SecretFlag
rs("AddTime")=now()
rs.update
rs.close
set rs=Nothing
Call SiteInfo()
If MesViewFlag = 0 Then
	response.write "<script language='javascript'>alert('留言成功提交，但系统已被设置为审核可见，请等待管理员审核、回复！');location.replace('MessageList.asp');</script>"
Else
	response.write "<script language='javascript'>alert('留言成功提交，请等待管理员回复！');location.replace('MessageList.asp');</script>"
End If
%>