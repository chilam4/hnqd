document.write("      <div class='r'>");
document.write("        <ul>");
document.write("          <li><a href='MemberRegister.asp'>会员注册</a></li>");
document.write("          <li> | </li>");
document.write("          <li><a href='MemberGetPass.asp'>找回密码</a></li>");
document.write("        </ul>");
document.write("      </div>");
document.write("      <form id='formLogin' name='formLogin' method='post' action='MemberLogin.asp'>");
document.write("        <div class='l'>");
<%If session("MemName")="" Then%>
document.write("          <input name='prev' type='hidden' id='prev' />");
document.write("          用户名：<input name='LoginName' type='text' id='LoginName' style='width:85px' class='memberName' />");
document.write("          密&nbsp;&nbsp;码：<input name='LoginPassword' type='password' id='LoginPassword' style='width:85px' class='memberName' />");
document.write("          <input name='submit' type='submit' value='登录' class='loginBt' />");
<% End If %>
<% If session("MemName")="sunly" Then %>
document.write("        您好：");
document.write("<%=session("MemName")%>"); 
document.write(" <a href='MemberCenter.Asp'>会员中心</a> <a href='ProductBuy.asp'>购物车</a> <a href='MemberLogin.asp?Action=Out'>退出登录</a>");
<% End If %>
document.write("        </div>");
document.write("      </form>");