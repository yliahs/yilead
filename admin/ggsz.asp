<!--#INCLUDE file="conn.asp"-->
<title>广告设置</title>
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
  sql="select fwxz from admin where name='"&Request.Cookies("admin")&"'"
  rs.open sql,conn,1,2
  
  if Request("action")="" then
%>
 <div class="p13">&nbsp;广告管理</div> 
  <form action="ggsz.asp?action=ad" method="post">
  <p>限制广告来源地址:
  <select name="fwxz" size="1">
   <%if rs("fwxz")="1" then%>
    <option value="1" selected="selected">开启</option>
    <option value="2">关闭</option>
    <%else%>
    <option value="1">开启</option>
    <option value="2" selected="selected">关闭</option>
    <%end if%>
    </select>
   </p> 
   <p>注:开启限制后,没有申请广告审核通过的站点将不能访问广告页面</p>

   
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Button1" value="确认设置" id="Button1"/><br/>
  </form>
<%
rs.close
set rs=nothing
end if%>


<%if Request("action")="ad" then
%><div class="p13">&nbsp;广告管理</div> <%
rs("fwxz")=Request("fwxz")
rs.update
rs.close
set rs=nothing
Response.write "<p>&nbsp;设置成功！</p>"
end if %>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

