<!--#INCLUDE file="conn.asp"-->
<title>个人中心</title>	
</head>
<body>
    <div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />个人中心</div>
<%set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from username where id="&Request("id")&""
rs.open sql,conn%>
帐号:<%=rs("username")%>&nbsp;<a href="xxuser.asp">下线推广</a><br/>金额：<%=rs("money")%><br/>Q&nbsp;Q：<%=rs("qq")%><br/>邮箱：<%=rs("email")%><br/>手机：<%=rs("sjh")%><br/>开户银行：<%=rs("khh")%><br/>开户地址：<%=rs("khdc")%><br/>开户姓名:<%=rs("khm")%><br/>银行帐号:<%=rs("yhzh")%><br/><a href="zzzl.asp">修改个人资料</a>
<p class="px"><img src="images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="index.asp">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>