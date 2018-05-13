<!--#INCLUDE file="conn.asp"-->
<title>查看公告</title>	
</head>
<body>
   <div id="all">     
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<%if Request.Cookies("username")="" then%>
<p><div class="p3"><a id="Top1_HyperLink1" href="Login.Asp">站长登录</a><span id="Top1_Label1">|</span><a id="Top1_HyperLink2" href="Reg.Asp">用户注册</a>&nbsp;</div></p>
<%else%>
<p><div class="p31">&nbsp;合作ID:<%=Request.Cookies("id")%>|<a href="pc.asp">个人中心</a>|<a href="sms.asp">信箱</a>|<a id="Top1_HyperLink1" href="login.asp?action=zx">注销</a></div></p>
<%end if%>
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * From wzgg where id="&Request("id")&""
  rs.open sql,conn,1,2
  if rs.eof then
  Response.Write "<p>此广告以删除或不存在！</p>"
  else
  rs("tj")=rs("tj")+1
  rs.update

%>
      <div class="p13"><img src="images/lalala.gif" width="14" height="14" />查看公告</div>
  <p><span id="Label_title" style="font-weight:bold;"><%=rs("title")%></span><br />
------------<br />
      <span id="Label_content">&nbsp;&nbsp;&nbsp;<%=rs("ggnl")%></span>
<br />
      <span id="Label_time">发布时间:<%=rs("time")%></span></p>
  <div></div>
  <%end if%>
  <%rs.close
  set rs=nothing%>
  <p><div class="px"><img src="images/fanhui.gif" width="16" height="9" /><a id="HyperLink1" href="index.Asp">返回上级</a></div>
  <div class="px"><img src="images/fanhui.gif" width="16" height="9" /><a href="/">返回首页</a></div></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
