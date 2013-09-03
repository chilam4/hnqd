<!--#include file="CheckAdmin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="images/index.css" type="text/css" rel="stylesheet" />
<link href="images/MasterPage.css" type="text/css" rel="stylesheet" />
<link href="images/Guide.css" type="text/css" rel="stylesheet" />
<title>管理导航菜单</title>

</head>

<body id="Guidebody">
  <div id="Guide_back">
    <ul>
      <li id="Guide_top">
        <div id="Guide_toptext">快捷导航</div>
      </li>
      <li id="Guide_main">
        <div id="Guide_box">
          <div class="guide">
            <ul id="Links">
<%
			dim id
			ID=request("id")
            if ID="System" or id="" then
				Response.Write "<li><a href=""SetSite.asp"" target=""main_right"">网站参数设置</a></li>"&vbCrlf
                Response.Write "<li><a href=""NavigationEdit.asp?Result=Add"" target=""main_right"">导航栏添加</a></li>"&vbCrlf
                Response.Write "<li><a href=""NavigationList.asp"" target=""main_right"">导航栏管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""FriendLinkEdit.asp?Result=Add"" target=""main_right"">友情链接添加</a></li>"&vbCrlf
                Response.Write "<li><a href=""FriendLinkList.asp"" target=""main_right"">友情链接管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""SetKey.asp"" target=""main_right"">站内链接管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""LinkEdit.asp?Result=Add"" target=""main_right"">站内链接添加</a></li>"&vbCrlf
                Response.Write "<li><a href=""SQL.asp"" target=""main_right"">数据库结构管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_SiteMap.asp"" target=""main_right"">生成谷歌SiteMap</a></li>"&vbCrlf
			end if
			if ID="News" then
				Response.Write "<li><a href=""NewsSort.asp?Action=Add&ParentID=0"" target=""main_right"">新闻类别管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""NewsList.asp"" target=""main_right"">新闻列表管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""NewsEdit.asp?Result=Add"" target=""main_right"">添加新闻</a></li>"&vbCrlf
                Response.Write "<li><a href=""AboutList.asp"" target=""main_right"">企业信息列表</a></li>"&vbCrlf
                Response.Write "<li><a href=""AboutEdit.asp?Result=Add"" target=""main_right"">添加企业信息</a></li>"&vbCrlf
				Response.Write "<li><a href=""ZhgkList.asp"" target=""main_right"">展会概况列表</a></li>"&vbCrlf
                Response.Write "<li><a href=""ZhgkEdit.asp?Result=Add"" target=""main_right"">添加展会概况</a></li>"&vbCrlf
				Response.Write "<li><a href=""ZsfwList.asp"" target=""main_right"">展商服务列表</a></li>"&vbCrlf
                Response.Write "<li><a href=""ZsfwEdit.asp?Result=Add"" target=""main_right"">添加展商服务</a></li>"&vbCrlf
			end if
			if ID="Product" then
				Response.Write "<li><a href=""ProductSort.asp?Action=Add&ParentID=0"" target=""main_right"">产品类别管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""ProductList.asp"" target=""main_right"">产品列表管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""ProductEdit.asp?Result=Add"" target=""main_right"">添加产品信息</a></li>"&vbCrlf
			end if
			if ID="DownLoad" then
				Response.Write "<li><a href=""DownSort.asp?Action=Add&ParentID=0"" target=""main_right"">下载类别管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""DownList.asp"" target=""main_right"">下载列表管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""DownEdit.asp?Result=Add"" target=""main_right"">添加下载信息</a></li>"&vbCrlf
			end if
			if ID="Talent" then 
				Response.Write "<li><a href=""JobsList.asp"" target=""main_right"">招聘列表管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""JobsEdit.asp?Result=Add"" target=""main_right"">添加招聘信息</a></li>"&vbCrlf
			end if
			if ID="Other" then
				Response.Write "<li><a href=""OthersSort.asp?Action=Add&ParentID=0"" target=""main_right"">信息类别管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""OthersList.asp"" target=""main_right"">信息列表管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""OthersEdit.asp?Result=Add"" target=""main_right"">添加信息</a></li>"&vbCrlf
			end if
			if ID="Feedback" then
				Response.Write "<li><a href=""MessageList.asp"" target=""main_right"">留言信息管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""OrderList.asp"" target=""main_right"">订单信息管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""TalentsList.asp"" target=""main_right"">人才信息管理</a></li>"&vbCrlf
			end if
			if ID="User" then
				Response.Write "<li><a href=""AdminList.asp"" target=""main_right"">网站管理员管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""AdminEdit.asp?Result=Add"" target=""main_right"">添加网站管理员</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemList.asp"" target=""main_right"">前台会员资料</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemGroup.asp"" target=""main_right"">会员组别管理</a></li>"&vbCrlf
                Response.Write "<li><a href=""MemGroup.asp?Result=Add"" target=""main_right"">添加会员组别</a></li>"&vbCrlf
                Response.Write "<li><a href=""ManageLog.asp"" target=""main_right"">后台登录日志管理</a></li>"&vbCrlf
			end if
			if ID="Html" then
				Response.Write "<li><a href=""Admin_htmlsort.asp"" target=""main_right"">生成产品分类页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlpro.asp"" target=""main_right"" onClick=""return Creathtml()"">生成所有产品详细页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlnewsort.asp"" target=""main_right"">生成新闻分类页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlnews.asp"" target=""main_right"" onClick=""return Creathtml()"">生成新闻详细页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlinfo.asp"" target=""main_right"">生成企业信息列表</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmldownsort.asp"" target=""main_right"">生成下载分类页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmldown.asp"" target=""main_right"" onClick=""return Creathtml()"">生成下载详细页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmljobsort.asp"" target=""main_right"">生成人才招聘列表</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmljob.asp"" target=""main_right"">生成人才招聘详细页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlothersort.asp"" target=""main_right"">生成其他信息分类页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlother.asp"" target=""main_right"" onClick=""return Creathtml()"">生成其他信息详细页面</a></li>"&vbCrlf
                Response.Write "<li><a href=""Admin_htmlindex.asp"" target=""main_right"">生成首页|询价|其他</a></li>"&vbCrlf
       			Response.Write "<li><a href=""Admin_html.asp"" target=""main_right"" onClick=""return CreathtmlAll()""><font color=""red"">生成全站静态页面</font></a></li>"&vbCrlf
			end if
			if ID="Plug" then
				Response.Write "<li><a href=""UserMessage.asp"" target=""main_right"">客户即时咨询管理</a></li>"&vbCrlf
			end if
			if ID="Promotion" then
				Response.Write "<li><a href=""http://www.baidu.com/search/url_submit.html"" target=""_blank"">百度登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.google.com/intl/zh-CN/add_url.html"" target=""_blank"">Google登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://search.help.cn.yahoo.com/h4_4.html"" target=""_blank"">Yahoo登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://search.msn.com/docs/submit.aspx"" target=""_blank"">Live登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.dmoz.org/World/Chinese_Simplified/"" target=""_blank"">Dmoz登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://www.alexa.com/site/help/webmasters"" target=""_blank"">Alexa登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://ads.zhongsou.com/register/page.jsp"" target=""_blank"">中搜登录入口</a></li>"&vbCrlf
                Response.Write "<li><a href=""http://iask.com/guest/add_url.php"" target=""_blank"">爱问登录入口</a></li>"&vbCrlf
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
    var bln=confirm("注意：添加、修改、删除相关数据时会自动生成、更新、删除所生成的静态文件。\n如果您没有对模板作过修改，不需要批量生成所有商品或新闻详细页面！\n如果您仅对产品、新闻、下载、人才等分类页面作过修改，只需要生成相关分类页面。\n\n请确定是否操作？");
    return bln;
}
function CreathtmlAll()
{
    var bln=confirm("警告：批量生成全站静态页面将耗费较多系统资源！\n\n请确定是否操作？");
    return bln;
}
-->
</script>