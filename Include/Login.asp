document.write("      <div class='r'>");
document.write("        <ul>");
document.write("          <li><a href='MemberRegister.asp'>��Աע��</a></li>");
document.write("          <li> | </li>");
document.write("          <li><a href='MemberGetPass.asp'>�һ�����</a></li>");
document.write("        </ul>");
document.write("      </div>");
document.write("      <form id='formLogin' name='formLogin' method='post' action='MemberLogin.asp'>");
document.write("        <div class='l'>");
<%If session("MemName")="" Then%>
document.write("          <input name='prev' type='hidden' id='prev' />");
document.write("          �û�����<input name='LoginName' type='text' id='LoginName' style='width:85px' class='memberName' />");
document.write("          ��&nbsp;&nbsp;�룺<input name='LoginPassword' type='password' id='LoginPassword' style='width:85px' class='memberName' />");
document.write("          <input name='submit' type='submit' value='��¼' class='loginBt' />");
<% End If %>
<% If session("MemName")="sunly" Then %>
document.write("        ���ã�");
document.write("<%=session("MemName")%>"); 
document.write(" <a href='MemberCenter.Asp'>��Ա����</a> <a href='ProductBuy.asp'>���ﳵ</a> <a href='MemberLogin.asp?Action=Out'>�˳���¼</a>");
<% End If %>
document.write("        </div>");
document.write("      </form>");