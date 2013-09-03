<!--#include file="Include/Const.asp"-->
<!--#include file="Include/ConnSiteData.asp"-->
<!--#include file="Include/Md5.asp"-->
<%
dim ID,mRealName,mSex,mPassword,vPassword,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,mHomePage
dim CheckCode
dim rs,sql
ID=request.QueryString("ID")
mRealName=trim(request.form("RealName"))
mSex=trim(request.form("Sex"))
mPassword=trim(request.form("Password"))
vPassword=trim(request.form("vPassword"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mTelephone=trim(request.form("Telephone"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
mHomePage=trim(request.form("HomePage"))
CheckCode = Trim(request.form("CheckCode"))
dim ErrMessage,ErrMsg(9),FindErr(9),i
  ErrMsg(0)="用户密码长度请保持在6-16位。"
  ErrMsg(1)="两次输入的用户密码不一致，请返回修改。"
  ErrMsg(2)="单位名称、详细地址长度请保持在100个字符以内。"
  ErrMsg(3)="请正确填写邮政编码。"
  ErrMsg(4)="请正确填写真实姓名、联系电话、传真号码、手机号码、网址。"
  ErrMsg(5)="电子信箱格式错误。"
  ErrMsg(6)="电子信箱已经注册过，请返回修改。"
  ErrMsg(7)="验证码不能为空，请返回检查。"
  ErrMsg(8)="您在【修改会员资料】页面停留的时间过长，导致验证码失效。\n请返回并刷新【修改会员资料】页面重新修改！"
  ErrMsg(9)="您输入的验证码和系统产生的不一致，请重新输入。"
if len(mPassword)>0 then
   if not (6<=len(mPassword) and len(mPassword)<=16) then FindErr(0)=true
   if mPassword<>vPassword then FindErr(1)=true
end if
if len(mCompany)=0 Or len(mCompany)>100 Or len(mAddress)=0 Or len(mAddress)>100 then FindErr(2)=true
if len(mZipCode)<>6 then FindErr(3)=true
if len(mRealName)=0 Or len(mTelephone)=0 Or len(mFax)=0 Or len(mMobile)=0 Or len(mHomePage)=0 Or len(mRealName)>50 Or len(mTelephone)>50 Or len(mFax)>50 Or len(mMobile)>50 Or len(mHomePage)>50 then FindErr(4)=true
if not IsValidEmail(mEmail) then FindErr(5)=true
if not conn.execute("select MemName from Qianbo_Members where ID<>"&ID&" and Email='" & mEmail & "'").eof then FindErr(6)=True
If CheckCode = "" Then FindErr(7)=true
If Trim(Session("CheckCode")) = "" Then FindErr(8)=true
If CheckCode <> Session("CheckCode") Then FindErr(9)=true
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
sql="select * from Qianbo_Members where ID="&ID
rs.open sql,conn,1,3
rs("RealName")=StrReplace(mRealName)
rs("Sex")=mSex
if len(mPassword)>0 then rs("Password")=Md5(mPassword)
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Telephone")=StrReplace(mTelephone)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("HomePage")=StrReplace(mHomePage)
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>alert('会员资料修改成功，即将返回到会员中心！');location.replace('MemberInfo.asp');</script>"
%>
