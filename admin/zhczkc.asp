<!--#INCLUDE file="conn.asp"-->
<title>帐户管理</title>
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
action=Request("action")
if action="cz" then
call cz
elseif action="kc" then
call kc
elseif action="czyz" then
call czyz
elseif action="kcyz" then
call kcyz
end if
%>

<%Function cz%>
<div class="p13">&nbsp;账户冲值</div>
<form action="zhczkc.asp?action=czyz" method="post">
<p>用户帐号：</p>
<p><input name="name" type="text" id="name" size="12" /></p>
<p>冲值金额：（单位：元）</p>
<p><input name="money" type="text" id="money" size="12" /></p>
<p><input name="ff" type="submit" id="ff" value="确认冲值" /></p>
</form>
<%End Function%>

<%Function czyz
Response.write "<div class='p13'>&nbsp;账户冲值</div>"
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from username where username='"&Request("name")&"'"
  rs.open sql,conn,1,2
  if not rs.eof then
rs("money")=rs("money")+Request("money")
rs.UPdate
Response.write "<p>往帐号："&Request("name")&"冲值"&Request("money")&"元成功！</p>"
'=====================财务写入日志
exec="select * from cwrz"
		set rsab=server.createobject("adodb.recordset")
		rsab.open exec,conn,1,2
		rsab.addnew
		rsab("username")=Request("name")
		rsab("money")=Request("money")
		rsab("sm")="往帐号："&Request("name")&"冲值"&Request("money")&"元成功！"
		rsab.update
		rsab.close
		set rsab=nothing
else
Response.write "<p>冲值失败，原因：找不到此用户！</p>"
end if
End Function%>



<%Function kc%>
<div class="p13">&nbsp;账户金额扣除</div>
<form action="zhczkc.asp?action=kcyz" method="post">
<p>用户帐号：</p>
<p><input name="name" type="text" id="name" size="12" /></p>
<p>扣除金额：（单位：元）</p>
<p><input name="money" type="text" id="money" size="12" /></p>
<p><input name="ff" type="submit" id="ff" value="确认扣除" /></p>
</form>
<%End Function%>

<%Function kcyz
Response.write "<div class='p13'>&nbsp;账户扣除</div>"
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from username where username='"&Request("name")&"'"
  rs.open sql,conn,1,2
  if not rs.eof then
rs("money")=rs("money")-Request("money")
rs.UPdate
Response.write "<p>帐号："&Request("name")&"扣除"&Request("money")&"元成功！</p>"
'=====================财务写入日志
exec="select * from cwrz"
		set rsab=server.createobject("adodb.recordset")
		rsab.open exec,conn,1,2
		rsab.addnew
		rsab("username")=Request("name")
		rsab("money")=Request("money")
		rsab("sm")="帐号："&Request("name")&"扣除"&Request("money")&"元成功！"
		rsab.update
		rsab.close
		set rsab=nothing
else
Response.write "<p>扣除失败，原因：找不到此用户！</p>"
end if
End Function%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
