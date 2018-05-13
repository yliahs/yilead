<!--#INCLUDE file="../conn.asp"-->
<!--#INCLUDE file="../hs.asp"-->
<!--#INCLUDE file="css.css"-->
<title>管理后台登陆中心</title>
</head>
<div id="all">

      <div class="p13"><img src="/images/lalala.gif" width="14" height="14" />管理员登录</div>
  <%
  dim action
  action=Request("action")
  if action="" then
  %>      
  <div align="center">
  <table width="80%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
    <tr>
      <td width="34%" align="center" style="height: 30px">帐&nbsp;&nbsp;&nbsp;号：</td>
	  <form action="login.asp?action=ad" method="post">
      <td width="66%" style="height: 30px">
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
          <input name="txt_check" type="text" size=15 maxlength=4 class="input"><img src="../checkcode.asp " alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='checkcode.asp?t='+(new Date().getTime());" >
    </tr>
  </table>
  <table width="70%" border="1" cellspacing="0" cellpadding="0" align="center" class="table1">
    <tr>
      <td height="25" align="center">
	  <input type="submit" name="Button1" value="登录" id="Button1" style="background-color:White;border:0" />&nbsp;&nbsp;<span style="color:#F00">|</span>&nbsp;&nbsp;
	  <input id="Reset1" type="reset" value="重新输入" style="border:0; background-color: #ffffff;" />
	  </form>
	  </td>
    </tr>
  </table>
  </div><br/>
  <%end if%>
  <%if action="ad" then%>
  <%
  dim user,pass
  user=Request("user")
  pass=HmacMd5(Request("pass"),2)
   if user="" then daving=daving&"帐号不可为空<br/>"
if pass="" then daving=daving&"密码不可为空<br/>"
if trim(session("validateCode")) <> trim(Request("txt_check")) then 
response.write("验证码错误，请重新输入<br/><a href='login.asp'>返回重写</a>")
response.end
end if 
if user="" or pass="" then 
     response.write "<card id='card1' title='出错提示!'>"
     response.write "<p align='left'>"
     response.write daving
     response.write "<br/><a href='login.asp'>返回重写</a>"
else

  set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where name='"&user&"' and pass='"&pass&"'"
  rsa.open sql,conn

  if rsa.eof and rsa.bof then
  Response.Write "<br/>用户名或密码错误!<br/>"
  response.write "<br/><a href='login.asp'>返回重写</a><br/><br/>"
  else
  Response.Cookies("admin").expires=now+1
   Response.Cookies("pass").expires=now+1

  Response.Cookies("admin")=rsa("name")
  Response.Cookies("pass")=rsa("pass")
Response.Redirect "index.asp"
  end if
  rsa.close
  set rsa=nothing
  end if
  

 %> 
  <%end if%>
  <%if action="zx" then
  Response.Cookies("admin")=""
  Response.Cookies("pass")=""
  Response.Write "<br/>&nbsp;注销成功!<br/>"
  Response.write "<p><a href='login.asp'>登陆中心</a></p>"
  end if%>
  
  

</div>
<body>
</body>
</html>

