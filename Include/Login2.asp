        <div class='r'>
    <ul>
      <li><a href='MemberRegister.asp'>��Աע��</a></li>
      <li> | </li>
      <li><a href='MemberGetPass.asp'>�һ�����</a></li>
    </ul>
  </div>
<form id='form' name='form' method='post' action='MemberLogin.asp'>
<% If session("MemName")="" Then %>
    <div class='l'>
      <input name='prev' type='hidden' id='prev' />
      �û�����<input name='LoginName' type='text' id='LoginName' style='width:85px' class='memberName' />
      ��&nbsp;&nbsp;�룺<input name='LoginPassword' type='password' id='LoginPassword' style='width:85px' class='memberName' />
      <input name='submit' type='submit' value='��¼' class='loginBt' />
    </div>
<% Else %>
<div class='l'>&nbsp;���ã�<%= session("MemName") %> <a href="MemberCenter.Asp">��Ա����</a> <a href="ProductBuy.asp">���ﳵ</a> <a href="MemberLogin.asp?Action=Out">�˳���¼</a></div>
<% End If %>
</form>