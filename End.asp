<div id="footer" style="border-top:1px dashed #ccc;">
	<p style="text-align:center;line-height:1.5em;">
        版权所有：湖南启达会展有限公司 地址：长沙市五一大道芙蓉华天大酒店东楼写字楼5038室 邮编：410001 <br/>
        经营许可证编号：湘ICP备12000461号-1
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