        <div class='r'>
    <ul>
      <li><a href='MemberRegister.asp'>会员注册</a></li>
      <li> | </li>
      <li><a href='MemberGetPass.asp'>找回密码</a></li>
    </ul>
  </div>
<form id='form' name='form' method='post' action='MemberLogin.asp'>
<% If session("MemName")="" Then %>
    <div class='l'>
      <input name='prev' type='hidden' id='prev' />
      用户名：<input name='LoginName' type='text' id='LoginName' style='width:85px' class='memberName' />
      密&nbsp;&nbsp;码：<input name='LoginPassword' type='password' id='LoginPassword' style='width:85px' class='memberName' />
      <input name='submit' type='submit' value='登录' class='loginBt' />
    </div>
<% Else %>
<div class='l'>&nbsp;您好：<%= session("MemName") %> <a href="MemberCenter.Asp">会员中心</a> <a href="ProductBuy.asp">购物车</a> <a href="MemberLogin.asp?Action=Out">退出登录</a></div>
<% End If %>
</form>