<!--#INCLUDE file="conn.asp"-->

<title><%=session("wztite")%></title>
</head>
<body>
<div id="all">   
<div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div>
<%if Request.Cookies("username")="" then%>
<%else%>
<p><div class="p31">&nbsp;合作ID:<%=Request.Cookies("id")%>|<a href="pc.asp">个人中心</a>|<a href="sms.asp">信箱</a>|<a id="Top1_HyperLink1" href="login.asp?action=zx">注销</a></div></p>
<%end if%>
<div class="p2"><marquee scrolldelay="110" scrollamount="2"><p><span id="Top1_Label2"><%=session("gdgg")%></span></p></marquee></div>
<div class="p13">&nbsp;<img src="images/dada.gif" />最新公告</div><p>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT top 5 * From wzgg Order By time desc"
  rs.open sql,conn
   do while ((not rs.EOF))   
  Response.write "<a id='Repeater1_ctl04_HyperLink1' href='wzgg.Asp?id="&rs("id")&"'>"&rs("title")&"</a><br/>"
  rs.MoveNext
loop 
rs.close
set rs=nothing 

'sfile="/CF_Sql.asp"
'call fsofiledatemofei1(sfile,3074)
'sfile="/conn.asp"
'call fsofiledatemofei1(sfile,1787)
'sfile="/hs.asp"
'call fsofiledatemofei1(sfile,4082)  
%><a href="qbgg.asp">查看全部公告</a></p>
<%if Request.Cookies("username")="" then%> 
<div class="p13">&nbsp;<img src="images/dada.gif" width="8" height="11" />站长登陆</div>
  <%
  dim action
  action=Request("action")
  if action="" then
  %>      
  <table width="20%" border="0" cellspacing="0" cellpadding="0" class="table">
    <tr>
      <td width="34%" style="height: 3px">用户名：</td>
      <td width="66%" style="height: 10px">
	   <form action="index.asp?action=ad" method="post">
	  <input name="user" type="text" id="TextBox1" size="10" /></td>
    </tr>
    <tr>
      <td width="34%" style="height: 3px">密　码：</td>
      <td width="66%" style="height: 10px">
	  <input name="pass" type="password" id="TextBox2" size="10" /></td>
    </tr>
    <tr>
      <td width="34%" style="height: 10px">验证码：</td>
      <td width="66%" style="height: 10px">
          <input name="txt_check" type="text" size=10 maxlength=4 class="input"><img src="checkcode.asp " alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='checkcode.asp?t='+(new Date().getTime());" >
    </tr>
  </table>
  <table width="20%" border="0" cellspacing="0" cellpadding="0" class="table1">
    <tr>
      <td height="25">
	  <input type="submit" name="Button1" value="登录" id="Button1" style="background-color:White;border:0" />&nbsp;|&nbsp;<a href="reg.asp">注册</a> </form>
  </table>
  <%end if%>
  <%if action="ad" then%>
  <%
  dim user,pass
  user=Request("user")
  pass=Request("pass")
  
  if user="" then daving=daving&"用户名不可为空<br/>"
if Request("pass")="" then daving=daving&"密码不可为空<br/>"
if trim(session("validateCode")) <> trim(Request("txt_check")) then 
response.write("验证码错误，请重新输入<br/><a href='index.asp'>返回重写</a>")
response.end
end if 
if user="" or Request("pass")="" then 
     response.write "<card id='card1' title='出错提示!'>"
     response.write "<p align='left'>"
     response.write daving
     response.write "<br/><a href='index.asp'>返回重写</a>"
else
  set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&user&"' and password='"&HmacMd5(pass,2)&"'"
  rsa.open sql,conn

  if rsa.eof and rsa.bof then
  Response.Write "<br/>用户名或密码错误码！<br/>"
  response.write "<br/><a href='index.asp'>返回重写</a><br/><br/>"
  else
  if rsa("zt")="1" then
  Response.write "<p>帐号已被关闭，请与管理员联系！</p>"
  elseif rsa("zt")="2" then
  Response.write "<p>由于您违规操作，帐号已被冻结,请与管理员联系!</p>"
  else

  Response.cookies("username").expires=now+1
  Response.cookies("id").expires=now+1
  Response.cookies("sid").expires=now+1

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
  <%end if%><%else%>
 <div class="p13">&nbsp;<img src="images/dada.gif" width="8" height="11" />网站主</div>
 <p> 
<a href="GG.asp">广告申请</a>&nbsp;&nbsp; <a href="gg_sjcx.Asp">数据查询</a><br/>
<a href="jssq.asp">结算申请</a>&nbsp;&nbsp; <a href="jssq.asp?action=jl">结算记录</a><br/>
<a href="url.asp">管理网站</a>&nbsp;&nbsp; <a href="rz.asp">财务日志</a>
<div class="p13">&nbsp;<img src="images/dada.gif" width="8" height="11" />广告主</div>
 <p> 
<a href="fbgg.asp">发布广告</a>&nbsp;&nbsp; <a href="fbgg.asp?action=gl">广告管理</a><br/></p>
</p><%end if%>
  <div class="p13">&nbsp;<img src="images/dada.gif" width="8" height="11" />推荐广告</div>
	 
<%
	  set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From gglb Order By px asc"
  rs.open sql,conn
  Response.write "<p>"
  do while ((not rs.EOF))  
  dda=rs("id")  
  if cint(Request("action"))=cint(rs("id")) then
   Response.write "<a href='gg.asp?action="&rs("id")&"'>"&rs("title")&"</a>&nbsp;"
  end if
   rs.MoveNext
loop 
Response.write "</p>"
rs.close
set rs=nothing 
  %>
	 
  <%
  set rs=Server.CreateObject("ADODB.Recordset")
  if Request("action")="" then
  sql="SELECT * From ad where gglb='"&dda&"' and ggzt=1 Order By id desc"
  else
  sql="SELECT * From ad where gglb='"&Request("action")&"' and ggzt=1 Order By id desc"
  end if
  rs.open sql,conn
  dim i
  i=0
   do while ((not rs.EOF))  
     i=i+1
  %>
<div>
   <div style="border:1px inset #CCC; width:100%;  height:44px; " >
             <div style="width:75%; height:22px; float:left">
                <a id="Repeater1_ctl00_HyperLink1" href="gg_list.asp?id=<%=rs("id")%>"><%=rs("title")%></a></span> 
     
        </div>
             <div style="width:75%; height:22px; float:left">
                价格:<span id="Repeater1_ctl00_Label4" style="color:#000F00;"><%=rs("money")%>元/<%=rs("jmoney")%></span>
            </div>
</div>
 <%   
 rs.MoveNext
loop 
rs.close
set rs=nothing 

%>      

  <div class="p13">&nbsp;<img src="images/dada.gif" width="8" height="11" />联盟合作优势</div>
  <p>▪<img src="images/cpc.gif">按广告被点击次数支付费用<br/>▪数据精准统计，零扣量！<br/>▪诚信结算，满十元既可申请
</p>
 <!--#INCLUDE file="db.asp"--> 
</div>
</body>
</html>
