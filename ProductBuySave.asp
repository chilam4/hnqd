<!--#include file="Include/Const.asp"-->
<!--#include file="Include/ConnSiteData.asp"-->
<%
if DateDiff("s",session("time"),now())<Refresh then
   response.write "<script language='JavaScript'>alert('防刷新机制启动：请不要在 "&Refresh&" 秒内重复刷新！');" & "history.back()" & "</script>"
   response.end
else
   session("time")=now()
end if
dim mOrderName,mRemark
dim mMemID,mRealName,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,CheckCode
dim rs,sql
mOrderName=trim(request.form("OrderName"))
mRemark="订购以下产品：<br />"&request.form("Products")&" 补充说明：<br />"&request.form("Remark")
mMemID=request.QueryString("MemberID")
mRealName=trim(request.form("RealName"))
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
  ErrMsg(0)="请填写订单标题。"
  ErrMsg(1)="请填写订购者姓名。"
  ErrMsg(2)="请填写单位名称、详细地址。"
  ErrMsg(3)="请正确填写邮政编码。"
  ErrMsg(4)="请正确填写联系电话(11个字符以上)。"
  ErrMsg(5)="请正确填写传真号码、手机号码。"
  ErrMsg(6)="电子信箱格式错误。"
  ErrMsg(7)="验证码不能为空，请返回检查。"
  ErrMsg(8)="您在【产品订购】页面停留的时间过长，导致验证码失效。\n请返回并刷新【产品订购】页面重新订购！"
  ErrMsg(9)="您输入的验证码和系统产生的不一致，请重新输入。"
if len(mOrderName)>100 Or len(mOrderName)=0 then
  FindErr(0)=true
end if
if len(mRealName)>50 Or len(mRealName)=0 then
  FindErr(1)=true
end if
if len(mCompany)>100 Or len(Address)>100 Or len(mCompany)=0 then
  FindErr(2)=true
end if
if len(mZipCode)<>6 then
  FindErr(3)=true
end if
if len(mTelephone)>50 Or len(mTelephone)<11 Or len(mTelephone)=0 then
  FindErr(4)=true
end if
if len(mFax)>50 Or len(mFax)=0 Or len(mMobile)=0 Or len(mMobile)>50 then
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
sql="select * from Qianbo_Order"
rs.open sql,conn,1,3
rs.addnew
rs("OrderName")=StrReplace(mOrderName)
rs("Remark")=StrReplace(mRemark)
rs("MemID")=mMemID
rs("Linkman")=StrReplace(mRealName)
rs("Sex")=mSex
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Telephone")=StrReplace(mTelephone)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("AddTime")=now()
rs.update
rs.close
set rs=nothing
Session("NoList")=""
response.write "<script language='javascript'>alert('订单提交成功！订单处理状态请登录会员中心查看！');location.replace('Index.asp');</script>"
%>