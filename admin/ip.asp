<!--#INCLUDE file="conn.asp"-->
<title>后台管理中心</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>

<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
 <div class="p13">&nbsp;IP清空</div> 
 <%if Request("action")="" then%>
 <p>确定清空IP吗！<br/><a href="ip.asp?action=ok">确定</a>&nbsp;&nbsp;<a href="index.asp">取消</a></p>
 <%else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from ip"
  conn.execute sql
 Response.write "<p>IP清空成功！</p>"
end if%>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 </div>
</body>
</html>

