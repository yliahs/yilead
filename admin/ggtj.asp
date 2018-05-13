<!--#INCLUDE file="conn.asp"-->
<title>广告浏览统计</title>
</head>
<body>
<div id="all">   

 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<div class="p13">&nbsp;广告统计</div>

 
 <%if Request("action")="" then
call index
elseif Request("action")="ok" then
call ok
end if
call waphx
 %>

<%sub index%>
<%set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ad Order By id desc"
  rs.open sql,conn,1,2
  if not rs.eof then
   i=0
 do while not rs.eof
 i=i+1
 Response.write "<p>"&i&".<a href='ggtj.asp?action=ok&amp;id="&rs("id")&"'>"&rs("title")&"</a></p>"
  rs.movenext
    	 loop
		 rs.close
		 set rs=nothing
		 end if%>
<%end sub%>


<%sub ok %>
 <form name="form1" method="post" action="ggtj.asp?action=ok&amp;id=<%=Request("id")%>">
日期： <input name="time1" type="text" value="<%=date-1%>" size="7" maxlength="20" />
到 <input name="time2" type="text" value="<%=date-1%>" size="7" maxlength="20"/>
<input type="submit" name="Button1" value="查询" /></form>
<%
time1=Request("time1")
time2=Request("time2")
if time1="" then
time1=date-1
end if
if time2="" then
time2=date-1
end if
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ggfw where ggid='"&Request("id")&"' and time>=#"&time1&"# and time<=#"&time2&"# Order By time desc"
rs.open sql,conn,1,2
Response.write "<p>日期"&time1&"至"&time2&"</p>"
if not rs.eof then
set rss=Server.CreateObject("ADODB.Recordset")
    sql="select * from ad where id="&rs("ggid")&""
rss.open sql,conn,1,2
if rss("gglx")=2 then
Response.write "<p>注：成功安装次数只针对jad软件有效<br/>独立下载只对下载广告有效<br/>有效注册只对注册广告有效<br/>-----------</p>"
end if
rss.close
set rss=nothing
i=0
 do while not rs.eof
 i=i+1
set rss=Server.CreateObject("ADODB.Recordset")
    sql="select * from ad where id="&rs("ggid")&""
rss.open sql,conn,1,2

Response.write "<p>"&i&".广告："&rs("title")&"<br/>点击IP："&rs("ip")&"<br/>点击PV："&rs("pv")&"<br/>"
if rss("gglx")=2 then
Response.write "独立下载："&rs("xzcs")&"次<br/>成功安装："&rs("anzcs")&"次<br/>有效注册："&rs("yxzc")&"个<br/>"
end if
Response.write "录入数据："&rs("tzsj")&"个<br/>支出费用："&rs("tzsj")&"x"&rss("money")&"="
if rs("tzsj")*rss("money")<1 then
Response.write "0"&rs("tzsj")*rss("money")&"元<br/>"
else
Response.write ""&rs("tzsj")*rss("money")&"元<br/>"
end if
Response.write "日期："&rs("time")&"<p>"
rss.close
set rss=nothing
rs.movenext
loop
else
Response.write "<p>暂无访问记录！</p>"
end if
rs.close
set rs=nothing
Response.Write "<p><img src='../images/fanhui.gif'/><a href='ggtj.asp'>返回上级</a></p>"
end sub%>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

