<!--#include file="CheckAdmin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="images/index.css" type="text/css" rel="stylesheet" />
<link href="images/MasterPage.css" type="text/css" rel="stylesheet" />
<link href="images/Guide.css" type="text/css" rel="stylesheet" />
<title>�������˵�</title>

</head>

<body id="Guidebody">
  <div id="Guide_back">
    <ul>
      <li id="Guide_top">
        <div id="Guide_toptext">��ݵ���</div>
      </li>
      <li id="Guide_main">
        <div id="Guide_box">
          <div class="guide">
            <ul id="Links">
<%
			dim id
			ID=request("id")
            if ID="System" or id="" then
				Response.Write "<li><a href=""SetSite.asp"" target=""main_right"">��վ��������</a></li>"&vbCrlf
                Response.Write "<li><a href=""NavigationEdit.asp?Result=Add"" target=""main_right"">���������</a></li>"&vbCrlf
                Response.Write "<li><a href=""NavigationList.asp"" target=""main_right"">����������</a></li>"&vbCrlf
                Response.Write "<li><a href=""FriendLinkEdit.asp?Result=Add"" target=""main_right"">�����������</a></li>"&vbCrlf
                Response.Write "<li><a href=""FriendLinkList.asp"" target=""main_right"">�������ӹ���</a></li>"&vbCrlf
                Response.Write "<li><a href=""SetKey.asp"" target=""main_right"">վ�����ӹ���</a></li>"&vbCrlf
                Response.Write "<li><a href=""LinkEdit.asp?Result=Add"" target=""main_right"">վ���������</a></li>"&vbCrlf
                Response.Write "<li><a href=""SQL.asp"" target=""main_right"">���ݿ�ṹ����</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_SiteMap.asp"" target=""main_right"">���ɹȸ�SiteMap</a></li>"&vbCrlf
			end if
			if ID="News" then
				Response.Write "<li><a href=""NewsSort.asp?Action=Add&ParentID=0"" target=""main_right"">����������</a></li>"&vbCrlf
                Response.Write "<li><a href=""NewsList.asp"" target=""main_right"">�����б����</a></li>"&vbCrlf
                Response.Write "<li><a href=""NewsEdit.asp?Result=Add"" target=""main_right"">�������</a></li>"&vbCrlf
                Response.Write "<li><a href=""AboutList.asp"" target=""main_right"">��ҵ��Ϣ�б�</a></li>"&vbCrlf
                Response.Write "<li><a href=""AboutEdit.asp?Result=Add"" target=""main_right"">�����ҵ��Ϣ</a></li>"&vbCrlf
				Response.Write "<li><a href=""ZhgkList.asp"" target=""main_right"">չ��ſ��б�</a></li>"&vbCrlf
                Response.Write "<li><a href=""ZhgkEdit.asp?Result=Add"" target=""main_right"">���չ��ſ�</a></li>"&vbCrlf
			end if
			if ID="Product" then
				Response.Write "<li><a href=""ProductSort.asp?Action=Add&ParentID=0"" target=""main_right"">��Ʒ������</a></li>"&vbCrlf
                Response.Write "<li><a href=""ProductList.asp"" target=""main_right"">��Ʒ�б����</a></li>"&vbCrlf
                Response.Write "<li><a href=""ProductEdit.asp?Result=Add"" target=""main_right"">��Ӳ�Ʒ��Ϣ</a></li>"&vbCrlf
			end if
			if ID="DownLoad" then
				Response.Write "<li><a href=""DownSort.asp?Action=Add&ParentID=0"" target=""main_right"">����������</a></li>"&vbCrlf
                Response.Write "<li><a href=""DownList.asp"" target=""main_right"">�����б����</a></li>"&vbCrlf
                Response.Write "<li><a href=""DownEdit.asp?Result=Add"" target=""main_right"">���������Ϣ</a></li>"&vbCrlf
			end if
			if ID="Talent" then 
				Response.Write "<li><a href=""JobsList.asp"" target=""main_right"">��Ƹ�б����</a></li>"&vbCrlf
                Response.Write "<li><a href=""JobsEdit.asp?Result=Add"" target=""main_right"">�����Ƹ��Ϣ</a></li>"&vbCrlf
			end if
			if ID="Other" then
				Response.Write "<li><a href=""OthersSort.asp?Action=Add&ParentID=0"" target=""main_right"">��Ϣ������</a></li>"&vbCrlf
                Response.Write "<li><a href=""OthersList.asp"" target=""main_right"">��Ϣ�б����</a></li>"&vbCrlf
                Response.Write "<li><a href=""OthersEdit.asp?Result=Add"" target=""main_right"">�����Ϣ</a></li>"&vbCrlf
			end if
			if ID="Feedback" then
				Response.Write "<li><a href=""MessageList.asp"" target=""main_right"">������Ϣ����</a></li>"&vbCrlf
                Response.Write "<li><a href=""OrderList.asp"" target=""main_right"">������Ϣ����</a></li>"&vbCrlf
                Response.Write "<li><a href=""TalentsList.asp"" target=""main_right"">�˲���Ϣ����</a></li>"&vbCrlf
			end if
			if ID="User" then
				Response.Write "<li><a href=""AdminList.asp"" target=""main_right"">��վ����Ա����</a></li>"&vbCrlf
                Response.Write "<li><a href=""AdminEdit.asp?Result=Add"" target=""main_right"">�����վ����Ա</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemList.asp"" target=""main_right"">ǰ̨��Ա����</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemGroup.asp"" target=""main_right"">��Ա������</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemGroup.asp?Result=Add"" target=""main_right"">��ӻ�Ա���</a></li>"&vbCrlf
                Response.Write "<li><a href=""ManageLog.asp"" target=""main_right"">��̨��¼��־����</a></li>"&vbCrlf
			end if
			if ID="Html" then
				Response.Write "<li><a href=""Admin_htmlsort.asp"" target=""main_right"">���ɲ�Ʒ����ҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlpro.asp"" target=""main_right"" onClick=""return Creathtml()"">�������в�Ʒ��ϸҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlnewsort.asp"" target=""main_right"">�������ŷ���ҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlnews.asp"" target=""main_right"" onClick=""return Creathtml()"">����������ϸҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlinfo.asp"" target=""main_right"">������ҵ��Ϣ�б�</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmldownsort.asp"" target=""main_right"">�������ط���ҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmldown.asp"" target=""main_right"" onClick=""return Creathtml()"">����������ϸҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmljobsort.asp"" target=""main_right"">�����˲���Ƹ�б�</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmljob.asp"" target=""main_right"">�����˲���Ƹ��ϸҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlothersort.asp"" target=""main_right"">����������Ϣ����ҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlother.asp"" target=""main_right"" onClick=""return Creathtml()"">����������Ϣ��ϸҳ��</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlindex.asp"" target=""main_right"">������ҳ|ѯ��|����</a></li>"&vbCrlf
       			Response.Write "<li><a href=""Admin_html.asp"" target=""main_right"" onClick=""return CreathtmlAll()""><font color=""red"">����ȫվ��̬ҳ��</font></a></li>"&vbCrlf
			end if
			if ID="Plug" then
				Response.Write "<li><a href=""UserMessage.asp"" target=""main_right"">�ͻ���ʱ��ѯ����</a></li>"&vbCrlf
			end if
			if ID="Promotion" then
				Response.Write "<li><a href=""http://www.baidu.com/search/url_submit.html"" target=""_blank"">�ٶȵ�¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.google.com/intl/zh-CN/add_url.html"" target=""_blank"">Google��¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://search.help.cn.yahoo.com/h4_4.html"" target=""_blank"">Yahoo��¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://search.msn.com/docs/submit.aspx"" target=""_blank"">Live��¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.dmoz.org/World/Chinese_Simplified/"" target=""_blank"">Dmoz��¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.alexa.com/site/help/webmasters"" target=""_blank"">Alexa��¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://ads.zhongsou.com/register/page.jsp"" target=""_blank"">���ѵ�¼���</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://iask.com/guest/add_url.php"" target=""_blank"">���ʵ�¼���</a></li>"&vbCrlf
			end if
%>
              
</ul>
          </div>
        </div>
      </li>
      <li id="Guide_bottom"></li>
    </ul>
  </div>
</body>
</html>
<script type="text/javascript">
<!--
function Creathtml()
{
    var bln=confirm("ע�⣺��ӡ��޸ġ�ɾ���������ʱ���Զ����ɡ����¡�ɾ�������ɵľ�̬�ļ���\n�����û�ж�ģ�������޸ģ�����Ҫ��������������Ʒ��������ϸҳ�棡\n��������Բ�Ʒ�����š����ء��˲ŵȷ���ҳ�������޸ģ�ֻ��Ҫ������ط���ҳ�档\n\n��ȷ���Ƿ������");
    return bln;
}
function CreathtmlAll()
{
    var bln=confirm("���棺��������ȫվ��̬ҳ�潫�ķѽ϶�ϵͳ��Դ��\n\n��ȷ���Ƿ������");
    return bln;
}
-->
</script>