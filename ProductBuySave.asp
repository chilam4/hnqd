<!--#include file="Include/Const.asp"-->
<!--#include file="Include/ConnSiteData.asp"-->
<%
if DateDiff("s",session("time"),now())<Refresh then
   response.write "<script language='JavaScript'>alert('��ˢ�»����������벻Ҫ�� "&Refresh&" �����ظ�ˢ�£�');" & "history.back()" & "</script>"
   response.end
else
   session("time")=now()
end if
dim mOrderName,mRemark
dim mMemID,mRealName,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,CheckCode
dim rs,sql
mOrderName=trim(request.form("OrderName"))
mRemark="�������²�Ʒ��<br />"&request.form("Products")&" ����˵����<br />"&request.form("Remark")
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
  ErrMsg(0)="����д�������⡣"
  ErrMsg(1)="����д������������"
  ErrMsg(2)="����д��λ���ơ���ϸ��ַ��"
  ErrMsg(3)="����ȷ��д�������롣"
  ErrMsg(4)="����ȷ��д��ϵ�绰(11���ַ�����)��"
  ErrMsg(5)="����ȷ��д������롢�ֻ����롣"
  ErrMsg(6)="���������ʽ����"
  ErrMsg(7)="��֤�벻��Ϊ�գ��뷵�ؼ�顣"
  ErrMsg(8)="���ڡ���Ʒ������ҳ��ͣ����ʱ�������������֤��ʧЧ��\n�뷵�ز�ˢ�¡���Ʒ������ҳ�����¶�����"
  ErrMsg(9)="���������֤���ϵͳ�����Ĳ�һ�£����������롣"
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
response.write "<script language='javascript'>alert('�����ύ�ɹ�����������״̬���¼��Ա���Ĳ鿴��');location.replace('Index.asp');</script>"
%>