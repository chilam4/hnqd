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
  ErrMsg(0)="�û����볤���뱣����6-16λ��"
  ErrMsg(1)="����������û����벻һ�£��뷵���޸ġ�"
  ErrMsg(2)="��λ���ơ���ϸ��ַ�����뱣����100���ַ����ڡ�"
  ErrMsg(3)="����ȷ��д�������롣"
  ErrMsg(4)="����ȷ��д��ʵ��������ϵ�绰��������롢�ֻ����롢��ַ��"
  ErrMsg(5)="���������ʽ����"
  ErrMsg(6)="���������Ѿ�ע������뷵���޸ġ�"
  ErrMsg(7)="��֤�벻��Ϊ�գ��뷵�ؼ�顣"
  ErrMsg(8)="���ڡ��޸Ļ�Ա���ϡ�ҳ��ͣ����ʱ�������������֤��ʧЧ��\n�뷵�ز�ˢ�¡��޸Ļ�Ա���ϡ�ҳ�������޸ģ�"
  ErrMsg(9)="���������֤���ϵͳ�����Ĳ�һ�£����������롣"
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
response.write "<script language='javascript'>alert('��Ա�����޸ĳɹ����������ص���Ա���ģ�');location.replace('MemberInfo.asp');</script>"
%>