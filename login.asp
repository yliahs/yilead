<!--#INCLUDE file="conn.asp"-->
<title>用户登陆</title>
</head>
<div id="all">
      
<div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div>
<div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div>

      <div class="p12"><img src="images/lalala.gif" width="14" height="14" />用户登录</div>
  <%
  dim action
  action=Request("action")
  if action="" then
  %>      
  <div align="center">
  <table width="80%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
    <tr>
      <td width="34%" align="center" style="height: 30px">用户名：</td>
      <td width="66%" style="height: 30px">
	   <form action="login.asp?action=ad" method="post">
	  <input name="user" type="text" id="TextBox1" size="15" /></td>
    </tr>
    <tr>
      <td width="34%" align="center" style="height: 30px">密　码：</td>
      <td width="66%" style="height: 30px">
	  <input name="pass" type="password" id="TextBox2" size="15" /></td>
    </tr>
    <tr>
      <td width="34%" align="center" style="height: 30px">验证码：</td>
      <td width="66%" style="height: 30px">
          <input name="txt_check" type="text" size=6 maxlength=4 class="input"><img src="checkcode.asp " alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='checkcode.asp?t='+(new Date().getTime());" >
    </tr>
  </table>
  <table width="70%" border="1" cellspacing="0" cellpadding="0" align="center" class="table1">
    <tr>
      <td height="25" align="center">
	  <input type="submit" name="Button1" value="登录" id="Button1" style="background-color:White;border:0" />&nbsp;&nbsp;<span style="color:#F00">|</span>&nbsp;&nbsp;
	  <input id="Reset1" type="reset" value="重新输入" style="border:0; background-color: #ffffff;" /> </form>
	 
	  </td>
    </tr>
  </table>
  </div><br/>
  <%end if%>
  <%if action="ad" then%>
  <%
  dim user,pass
  user=Request("user")
  pass=Request("pass")
  
  if user="" then daving=daving&"用户名不可为空<br/>"
if Request("pass")="" then daving=daving&"密码不可为空<br/>"
if trim(session("validateCode")) <> trim(Request("txt_check")) then 
response.write("验证码错误，请重新输入<br/><a href='login.asp'>返回重写</a>")
response.end
end if 
if user="" or Request("pass")="" then 
     response.write "<card id='card1' title='出错提示!'>"
     response.write "<p align='left'>"
     response.write daving
     response.write "<br/><a href='login.asp'>返回重写</a>"
else
  set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&user&"' and password='"&HmacMd5(pass,2)&"'"
  rsa.open sql,conn

  if rsa.eof and rsa.bof then
  Response.Write "<br/>用户名或密码错误码！<br/>"
  response.write "<br/><a href='login.asp'>返回重写</a><br/><br/>"
  else
  if rsa("zt")="1" then
  Response.write "<p>帐号已被关闭，请与管理员联系！</p>"
  elseif rsa("zt")="2" then
  Response.write "<p>由于您违规操作，帐号已被冻结,请与管理员联系!</p>"
  else
  
  Response.cookies("username").expires=now+1
  Response.cookies("id").expires=now+1
  Response.cookies("password").expires=now+1

  Response.Cookies("username")=rsa("username")
  Response.Cookies("password")=rsa("password")
  Response.Cookies("id")=rsa("id")
Response.Redirect "/"
end if
  end if
  rsa.close
  set rsa=nothing
  end if
 %> 
  <%end if%>
  <%if action="zx" then
  Response.Cookies("username")=""
  Response.Cookies("password")=""
  Response.Write "<p>注销成功!</p>"
  end if%>
  
  
  
  <div class="px"><img src="images/fanhui.gif" width="16" height="9" alt="."/><a href="Reg.asp">注册新用户</a></div>
<img src="images/fanhui.gif" width="16" height="9" alt="."/><a href="/">返回首页</a>

<!--#INCLUDE file="db.asp"-->

</div>
<body>
</body>
</html>
