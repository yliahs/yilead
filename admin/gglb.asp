<!--#INCLUDE file="conn.asp"-->
<title>广告分类</title>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--广告分类页面-->
<%
  if Request("action")="" then%>
  <div class="p13">&nbsp;广告分类管理</div>
   <% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb Order By px asc"
  rs.open sql,conn,1,2 
  dim i
  i=0
   do while ((not rs.eof))%>
   <%i=i+1%>
   <form name="form1" method="post" action="gglb.asp?action=ok">
<p><%=rs("title")%>&nbsp;&nbsp;排序：
<input name="title<%= i %>" type="text" size="2" value="<%=rs("px")%>" maxlength="5" />
&nbsp;<a href="gglb.asp?action=bj&amp;id=<%=rs("id")%>">编辑</a>&nbsp;&nbsp;<a href="gglb.asp?action=sc&amp;id=<%=rs("id")%>">删除</a>
<input name="id<%= i %>" type="hidden" value="<%=rs("id")%>">
  <%rs.MoveNext
loop 
rs.close
set rs=nothing%>  
<input name="ii" type="hidden" value="<%=i%>">
<p><input type="submit" name="Button1" value="确认编辑排序" /></form></p>
<%end if%>

<%if Request("action")="ok" then%>
<div class="p13">&nbsp;广告分类管理</div>
<%dim ii
ii=Request("ii")

for i=1 to ii
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb where id="&Request("id"&i)&""
  rs.open sql,conn,1,2 

if request("title"&i)="" then
Response.write "<p>所有内容不能为空！</p>"
else
rs("px")=Request("title"&i)
rs.update

end if
next
Response.write "<p>排序修改成功！</p>"
rs.close
set rs=nothing
%>
<%end if%>

<%if Request("action")="tj" then%>
<div class="p13">&nbsp;添加分类</div>
<%if Request("ac")="" then%>
<form name="form1" method="post" action="gglb.asp?action=tj&amp;ac=ok">
<p>分类名称：</p>
<input name="title" type="text" value="" />
<p>排序号(数字,最小排在最前面)：</p>
<input name="px" type="text" value="" maxlength="5" />
<p><input type="submit" name="Button1" value="确认添加" /></form></p>
<%else%>
<%
if Request("title")="" or Request("px")="" then
Response.write "<p>分类名称或排序号不能为空！</p>"
else


set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb"
  rs.open sql,conn,1,2 
rs.addnew
rs("title")=Request("title")
rs("px")=Request("px")
rs.update
rs.close
set rs=nothing
Response.Write "<p>广告分类添加成功！</p>"
end if
%>
<%end if%>
<%end if%>


<%if Request("action")="bj" then%>
<div class="p13">&nbsp;编辑广告分类</div>
<%set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb where id="&Request("id")&""
  rs.open sql,conn,1,2 
  if rs.eof then
  Response.write "<p>此广告分类不存在！</p>"
  end if
  if Request("ac")="" then
%>
<form action="gglb.asp?action=bj&amp;ac=ok&amp;id=<%=Request("id")%>" method="post">
<p>分类：<input name="name" type="text" value="<%=rs("title")%>"><br/>
排序：<input name="px" type="text" value="<%=rs("px")%>"><br/>
<input name="" type="submit" value="确定编辑">
</p>
</form>
<%else
rs("title")=Request("name")
rs("px")=Request("px")
rs.update
rs.close
set rs=nothing
Response.write "<p>广告分类编辑成功！</p>"
end if%>
<%end if%>


<%if Request("action")="sc" then%>
<div class="p13">&nbsp;删除广告分类</div>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb where id="&Request("id")&""
  rs.open sql,conn,1,2 
  if rs.eof then
  Response.write "<p>此广告分类不存在！</p>"
  end if
  
if Request("ac")="" then
Response.write "<p>确定删除"""&rs("title")&"""分类吗？<br/><a href='gglb.asp?action=sc&amp;ac=ok&amp;id="&Request("id")&"'>确定删除</a></p>"
else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from gglb where id="&Request("id")&""
  conn.execute sql

Response.write "<P>分类删除成功</p>"
end if
rs.close
set rs=nothing
end if%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
