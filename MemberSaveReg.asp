<!--#include file="Include/Const.asp"-->
<!--#include file="Include/ConnSiteData.asp"-->
<!--#include file="Include/Md5.asp"-->
<%
if DateDiff("s",session("time"),now())<Refresh then
   response.write "<script language='JavaScript'>alert('��ˢ�»����������벻Ҫ�� "&Refresh&" �����ظ�ˢ�£�');" & "history.back()" & "</script>"
   response.end
else
   session("time")=now()
end if
dim mMemName,mRealName,mSex,mPassword,mQuestion,mAnswer,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,mHomePage
dim vPassword,CheckCode
dim rs,sql
mMemName=trim(request.form("MemName"))
mRealName=trim(request.form("RealName"))
mSex=trim(request.form("Sex"))
mPassword=trim(request.form("Password"))
vPassword=trim(request.form("vPassword"))
mQuestion=trim(request.form("Question"))
mAnswer=trim(request.form("Answer"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mTelephone=trim(request.form("Telephone"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
mHomePage=trim(request.form("HomePage"))
CheckCode = Trim(request.form("CheckCode"))
dim ErrMessage,ErrMsg(13),FindErr(13),i
  ErrMsg(0)="�û�������Ϊ0-9,a-z,-_������ϵ�3-16���ַ���"
  ErrMsg(1)="�û����ظ����뷵���޸ġ�"
  ErrMsg(2)="�û����볤���뱣����6-16λ��"
  ErrMsg(3)="����������û����벻һ�£��뷵���޸ġ�"
  ErrMsg(4)="���뱣�����ⳤ���뱣����3-100���κ��ַ���"
  ErrMsg(5)="���뱣���𰸳����뱣����3-100���κ��ַ���"
  ErrMsg(6)="��λ���ơ���ϸ��ַ�����뱣����100���ַ����ڡ�"
  ErrMsg(7)="����ȷ��д�������롣"
  ErrMsg(8)="����ȷ��д��ʵ��������ϵ�绰��������롢�ֻ����롢��ַ��"
  ErrMsg(9)="���������ʽ����"
  ErrMsg(10)="���������Ѿ�ע������뷵���޸ġ�"
  ErrMsg(11)="��֤�벻��Ϊ�գ��뷵�ؼ�顣"
  ErrMsg(12)="���ڡ���Աע�᡿ҳ��ͣ����ʱ�������������֤��ʧЧ��\n�뷵�ز�ˢ�¡���Աע�᡿ҳ������ע�ᣡ"
  ErrMsg(13)="���������֤���ϵͳ�����Ĳ�һ�£����������롣"
if not IsValidMemName(mMemName) then FindErr(0)=true
if not conn.execute("select MemName from Qianbo_Members where MemName='" & mMemName & "'").eof then FindErr(1)=true
if not (6<=len(mPassword) and len(mPassword)<=16) then FindErr(2)=true
if mPassword<>vPassword then FindErr(3)=true
if not (3<=len(mQuestion) and len(mQuestion)<=100) then FindErr(4)=true
if not (3<=len(mAnswer) and len(mAnswer)<=100) then FindErr(5)=true
if len(mCompany)=0 Or len(mCompany)>100 Or len(mAddress)=0 Or len(mAddress)>100 then FindErr(6)=true
if len(mZipCode)<>6 then FindErr(7)=true
if len(mRealName)=0 Or len(mTelephone)=0 Or len(mFax)=0 Or len(mMobile)=0 Or len(mHomePage)=0 Or len(mRealName)>50 Or len(mTelephone)>50 Or len(mFax)>50 Or len(mMobile)>50 Or len(mHomePage)>50 then FindErr(8)=true
if not IsValidEmail(mEmail) then FindErr(9)=true
if not conn.execute("select MemName from Qianbo_Members where Email='" & Email & "'").eof then FindErr(10)=True
If CheckCode = "" Then FindErr(11)=true
If Trim(Session("CheckCode")) = "" Then FindErr(12)=true
If CheckCode <> Session("CheckCode") Then FindErr(13)=true
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
sql="select * from Qianbo_Members"
rs.open sql,conn,1,3
rs.addnew
rs("MemName")=mMemName
rs("RealName")=StrReplace(mRealName)
rs("Sex")=mSex
rs("Password")=Md5(mPassword)
rs("Question")=StrReplace(mQuestion)
rs("Answer")=Md5(mAnswer)
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Telephone")=StrReplace(mTelephone)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("HomePage")=StrReplace(mHomePage)
rs("GroupID")="200709088888888888"
rs("GroupName")=GroupName("200709088888888888")
rs("AddTime")=now()
rs.update
rs.close
Set rs=Nothing
response.write "<script language='javascript'>alert('ע��ɹ������¼��');location.replace('Index.asp');</script>"
response.End

function GroupName(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupName=rs("GroupName")
  rs.close
  set rs=nothing
end function
%>