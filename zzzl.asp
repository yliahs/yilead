<!--#INCLUDE file="conn.asp"-->
<title>资料管理</title>	

</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

<%if Request("action")="" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />资料管理</div>
<p><a href="zzzl.asp?action=xgmm">修改登陆密码</a></p>
<p><a href="zzzl.asp?action=yhzl">修改银行资料</a></p>
<p><a href="zzzl.asp?action=xx">修改个人信息</a></p>
<p><a href="up/index.asp">修改证件信息</a></p>
<%end if%>


<%if Request("action")="xgmm" then%>
<%if Request("ac")="" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />修改登陆密码</div>
 <form name="form1" method="post" action="zzzl.asp?action=xgmm&amp;ac=ok">
<p>密码:(英文或数字6-15位)</p>
  <input name="password" type="text" maxlength="15" id="TextBox1" style="width:120px;" />
<p>重复输入密码:</p>
   <p><input name="password1" type="text" maxlength="15" id="TextBox1" style="width:120px;" /></p>
   <p><input type="submit" name="Button1" value="确认修改" /></p>
	</form>
	<br/>
<img src="images/fanhui.gif" width="16" height="9"/><a href="zzzl.asp">资料管理</a>
	<%
	else
	%><div class="p12"> <img src="images/lalala.gif" width="14" height="14" />修改登陆密码</div><br/><%
	dim password,password1
	password=Request("password")
	password1=Request("password1")
	if password<>password1 then
	Response.write "两次输入的密码不一致！<br/><br/><a href='zzzl.asp?action=xgmm'>返回重新输入</a><br/>"
	else
	 Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT password From username where username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
rs("password")=HmacMd5(password,2)
rs.Update
rs.close
set rs=nothing
Response.write "<br/>密码修改成功！<br/><br/><img src='images/fanhui.gif' width='16' height='9'/><a href='zzzl.asp'>资料管理</a><br/>"
	end if
	end if%>
<%end if%>


<%if Request("action")="yhzl" then 
if Request("ac")="" then
Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT khh,khdc,khm,yhzh From username where username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />修改银行资料</div>
 <form name="form1" method="post" action="zzzl.asp?action=yhzl&amp;ac=ok">
 <p>开户银行：</p>
 <select name="khh" id="DropDownList1">
		<option selected="selected" value="工商银行">工商银行</option>
		<option value="支付宝">支付宝</option>
		<option value="财付通">财付通</option>
	</select>
 <p>收款帐号:</p>
  <input name="yhzh" type="text" id="TextBox1" style="width:120px;" value="<%=rs("yhzh")%>" maxlength="20" />
 <p>开户地址:(支付宝、财付通可不填)</p>
   <input name="khdc" type="text" id="TextBox1" style="width:120px;" value="<%=rs("khdc")%>" maxlength="50" />
   <p>开户名:</p>
    <input name="khm" type="text" id="TextBox1" style="width:120px;" value="<%=rs("khm")%>" maxlength="10" />
    <p><input type="submit" name="Button1" value="确认修改" /></p>
  </form>
  <img src="images/fanhui.gif" width="16" height="9"/><a href="zzzl.asp">资料管理</a>
<%else
dim khh,yhzh,khdc,khm
khh=Request("khh")
yhzh=Request("yhzh")
khdc=Request("khdc")
khm=Request("khm")
if yhzh="" then avv=avv&"<br/>收款帐号不可为空!<br/>"
if khh<>"支付宝" then
if khm="" then avv=avv&"<br/>开户名不可为空！<br/>"
if khdc="" then avv=avv&"<br/>开户地址不可为空<br/>"
end if
if khh<>"支付宝" then
if yhzh="" or khm="" or khdc="" then
Response.write "<div class='p12'> <img src='images/lalala.gif' width='14' height='14' />修改银行资料</div>"
Response.write avv
Response.write "<br/><a href='zzzl.asp?action=yhzl'>返回重新输入</a><br/>"
Response.end
end if
else
if yhzh="" then
Response.write "<div class='p12'> <img src='images/lalala.gif' width='14' height='14' />修改银行资料</div>"
Response.write avv
Response.write "<br/><a href='zzzl.asp?action=yhzl'>返回重新输入</a><br/>"
Response.end
end if
end if
Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From username where username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
rs("khh")=khh
rs("khdc")=khdc
rs("khm")=khm
rs("yhzh")=yhzh
rs.Update
rs.close
set rs=nothing
Response.write "<br/>银行资料修改成功！<br/><br/><img src='images/fanhui.gif' width='16' height='9'/><a href='zzzl.asp'>资料管理</a><br/>"
end if
end if%>
  
 
 
<%if Request("action")="xx" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />修改个人信息</div>
<%if Request("ac")="" then%>
<%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT email,qq,sjh From username where username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
%>
<form name="form1" method="post" action="zzzl.asp?action=xx&amp;ac=ok">
  <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
  <tr>
		<td width="28%" height="22" align="center">
          <span id="Label1" style="color:#000F00;">电子邮箱</span> </td>
		<td width="72%" height="22">&nbsp;<span id="Label_title" style="color:#000F00;">
		  <input name="email" type="text" id="khm" style="width:120px;" value="<%=rs("email")%>" maxlength="25" />
		</span></td>
	</tr>
	<tr>
		<td width="28%" height="22" align="center">
          <span id="Label1" style="color:#000F00;">联系QQ</span> </td>
		<td width="72%" height="22">&nbsp;<span id="Label_title" style="color:#000F00;"> <input name="qq" type="text" id="TextBox1" style="width:120px;" value="<%=rs("qq")%>" maxlength="13" />
		</span></td>
	</tr>
	<tr>
		<td width="22%" height="22" align="center">
          <span id="Label1" style="color:#000F00;">联系电话</span> </td>
		<td width="78%" height="22">&nbsp;<span id="Label_title" style="color:#000F00;"> <input name="sjh" type="text" id="TextBox1" style="width:120px;" value="<%=rs("sjh")%>" maxlength="11" />
		</span></td>
	</tr>
	<tr>
		<td height="22" colspan="2" align="center">
	<input type="submit" name="Button1" value="确认修改" /></td></tr>
  </form></table>
  <br/> <img src="images/fanhui.gif" width="16" height="9"/><a href="zzzl.asp">资料管理</a>
  <%rs.close
  set rs=nothing%>
	<%else%>
	<%
	Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT email,qq,sjh From username where username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
	rs("email")=Request("email")
	rs("qq")=Request("qq")
	rs("sjh")=Request("sjh")
	rs.Update
	rs.close
	set rs=nothing
	Response.write "<br/>资料修改成功!<br/><br/><img src='images/fanhui.gif' width='16' height='9'/><a href='zzzl.asp'>资料管理</a><br/>"	
	%>
	<%end if%>
	<%end if%>
	
  
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>