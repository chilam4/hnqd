<div id="footer" style="border-top:1px dashed #ccc;">
	<p style="text-align:center;line-height:1.5em;">
        ��Ȩ���У����������չ���޹�˾ ��ַ����ɳ����һ���ܽ�ػ����Ƶ궫¥д��¥5038�� �ʱࣺ410001 <br/>
        ��Ӫ���֤��ţ���ICP��12000461��-1
    </p>
</div>
</body>
</html>
<%
function MemGroup(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from Qianbo_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  MemGroup=rs("GroupName")
  rs.close
  set rs=nothing
end Function
%>