<!--#INCLUDE file="conn.asp"-->
<title>管理帐号设置</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select name,pass from admin where name='"&Request.Cookies("admin")&"'"
  rs.open sql,conn,1,2
  
  if Request("action")="" then
%>
 <div class="p13">&nbsp;管理员设置</div> 
  <form action="admin.asp?action=ad" method="post">
  <p>管理员:</p> 
    <input name="name" type="text" id="TextBox2" style="border-color:Yellow;width:80%;" value="<%=rs("name")%>" size="15" />
   <p>密&nbsp;码:</p> 
   <input name="pass" type="text" id="TextBox2" style="border-color:Yellow;width:80%;" value="<%=rs("pass")%>" size="15" />
 
<br/><input type="submit" name="Button1" value="确认设置" id="Button1"/><br/>
  </form>
<%
rs.close
set rs=nothing
end if%>


<%if Request("action")="ad" then%>
<div class="p13">&nbsp;管理员设置</div> 
<%

rs("name")=Request("name")
rs("pass")=HmacMd5(Request("pass"),2)
rs.UPdate
rs.close
set rs=nothing
Response.write "<p>&nbsp;&nbsp;设置成功!</p>"
end if %>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

