<!--#INCLUDE file="conn.asp"-->
<title>公告管理</title>
<style type="text/css">
<!--
#all aaa {
	color: #00F;
}

-->
</style></head>
<body>
<div id="all">   
    <!--公告管理首页-->
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<%
  if Request("action")="" then
  
    for i=1 to 2
randomize
m=m&int((9)*rnd+1)
Next
  
  ada=5
  if request("page")<>"" then 
  page=cint(request.QueryString("page"))
  if page="" or page<=0 then page=1
  Else
  page=1
  End if
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from wzgg Order By time desc"
  rs.open sql,conn,1,2

   

%>
 <div class="p13">&nbsp;公告管理</div> 
 <p><a href="wzgg.asp?action=fb">发布公告</a></p>
 
 <%  if rs.eof then 
 Response.write "<p>暂无公告！</p>"
 Response.end
 end if
  rs.Move((page-1)*ada) 
  dim i
  i=1
 do while ((not rs.EOF) and i<=ada)%>
<p> <a href="wzgg.asp?id=<%=rs("id")%>&amp;action=ck"><%=rs("title")%></a></p>
<p><a href="wzgg.asp?id=<%=rs("id")%>&amp;action=ck">查看</a>&nbsp;<a href="wzgg.asp?id=<%=rs("id")%>&amp;action=xg">修改</a>&nbsp;<a href="wzgg.asp?id=<%=rs("id")%>&amp;action=sc">删除</a>&nbsp;阅:<%=rs("tj")%>次<br/></p>
 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
 <tr></tr>
 </table>
 
<%
rs.MoveNext
loop 
i=i+1 

if page*ada<rs.recordcount then
%><a href="wzgg.asp?page=<%=page+1%>">下一页</a><%
end if
if page>1 then 
%><a href="wzgg.asp?page=<%=page-1%>">上一页</a><%
end if
if rs.recordcount > 0 then
response.write("["& page & "/" & (int(rs.recordcount/ada)+1) &"页]")
end if 
rs.close
set rs=nothing
end if%>

<!--查看公告内容页面 -->
<%if Request("action")="ck" then%>
 <div class="p13">&nbsp;公告管理</div> 
<%set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from wzgg where id="&Request("id")&""
  rs.open sql,conn,1,2
  Response.write "<p><b>"&rs("title")&"</b></p>------------<br/>"
  Response.write "<p>&nbsp;&nbsp;&nbsp;"&rs("ggnl")&"</p>" 
  Response.write "<p>发布时间"&rs("time")&"</p>"
  rs.close
  set rs=nothing
Response.write "<p><img src='../images/fanhui.gif'/><a href='wzgg.asp'>公告管理</a></p>"  
end if %>

<!--修改公告内容页面-->
<%if Request("action")="xg" then
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from wzgg where id="&Request("id")&""
  rs.open sql,conn,1,2
  if Request("ac")="" then
  %>
   <div class="p13">&nbsp;公告修改</div> 
  <form name="form1" method="post" action="wzgg.asp?id=<%=Request("id")%>&amp;action=xg&amp;ac=ok">
  <p>公告标题:</p>
 <input name="title" type="text" id="TextBox1" value="<%=rs("title")%>" maxlength="20" />
 <p>公告内容:</p>
 <textarea name="ggnl" rows="4" id="TextBox1"><%=rs("ggnl")%></textarea>
<p>&nbsp;&nbsp;<input type="submit" name="Button1" value="确认修改" /></p>
  </form> 
 <p><img src="../images/fanhui.gif"/><a href="wzgg.asp">公告管理</a></p> 
<%end if%>
<% if Request("ac")="ok" then%>
<div class="p13">&nbsp;公告修改</div> 
<%
rs("title")=Request("title")
rs("ggnl")=Request("ggnl")
rs.UPdate
rs.close
set rs=nothing
Response.write "<p>公告修改成功!</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='wzgg.asp'>公告管理</a></p>"   
 end if
end if%>


<!--删除公告页面-->
<%if Request("action")="sc" then
set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from wzgg where id="&Request("id")&""
  conn.execute sql
  Response.write " <div class='p13'>&nbsp;公告管理</div> "
  Response.write "<p>公告删除成功!</p>"
  Response.write "<p><img src='../images/fanhui.gif'/><a href='wzgg.asp'>公告管理</a></p>"   
end if%>

<% if Request("action")="fb" then %>
<div class="p13">&nbsp;发布公告</div> 
<%if Request("ac")="" then%>
  <form name="form1" method="post" action="wzgg.asp?action=fb&amp;ac=ok">
  <p>公告标题:</p>
 <input name="title" type="text" id="TextBox1" value="" maxlength="20" />
 <p>公告内容:</p>
 <textarea name="ggnl" rows="4" id="TextBox1"></textarea>
<p>&nbsp;&nbsp;<input type="submit" name="Button1" value="确认发布" /></p>
  </form> 
  <%end if
  if Request("ac")="ok" then
  if Request("title")<>"" or Request("ggnl")<>"" then
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from wzgg"
  rs.open sql,conn,1,2  
  rs.addnew
  rs("title")=Request("title")
  rs("ggnl")=Request("ggnl")
  rs.update
  rs.close
  set rs=nothing
  end if
  Response.write "<p>公告发布成功!</p>"
   Response.write "<p><img src='../images/fanhui.gif'/><a href='wzgg.asp'>公告管理</a></p>"   
 end if%>
<%end if%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

