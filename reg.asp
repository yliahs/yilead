<!--#INCLUDE file="conn.asp"-->
<title>新用户注册</title>	
</head>
<body>
<div id="all">

<div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div>
<div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div>

   <div class="p12"><img src="images/lalala.gif" width="14" height="14" />新用户注册</div>
   <%if Request("action")="" then%>
    <form name="form1" method="post" action="Reg.asp?action=ok&amp;uid=<%=Request("uid")%>" onSubmit="javascript:return WebForm_OnSubmit();" id="form1">

      说明：用户名和密码只能用字母、数字或组合。所有项目都不要使用特殊字符，带*号为必填项.  
 <p>*用户名(英文或数字3-15位):<br/>
       <input name="user" type="text" maxlength="15" id="TextBox1" style="width:120px;" /><br/>
     *密码(英文或数字6-15位)<br/>
    <input name="password" type="text" maxlength="15" id="TextBox2" style="width:120px;" /><br/>
  *联系手机:<br/>
  <input name="sjh" type="text" maxlength="11" id="TextBox3" style="width:120px;" /><br/>

    *QQ:<br/>
   <input name="qq" type="text" maxlength="12" id="TextBox5" style="width:120px;" /><br/>

    手机号和QQ请填写正确，客服将联系审核<br/>

 *网站名称: <br/>

  <input name="urltitle" type="text" maxlength="10" id="TextBox6" style="width:120px;" /><br/>

   *网站分类:<br/>

 <select name="urllx" id="DropDownList1">
		<option selected="selected" value="门户">门户</option>
		<option value="社区">社区</option>
		<option value="网址">网址</option>
		<option value="下载">下载</option>
		<option value="书刊">书刊</option>
		<option value="两性">两性</option>
		<option value="彩票">彩票</option>
		<option value="行业">行业</option>
		<option value="企业">企业</option>
		<option value="个人博客">个人博客</option>
		<option value="资讯">资讯</option>
		<option value="动漫">动漫</option>
		<option value="汽车">汽车</option>
		<option value="手机">手机</option>
		<option value="图片">图片</option>
		<option value="其它">其它</option>

	</select><br/>
    *网址:<br/>
  <input name="url" type="text" value="http://" id="TextBox7" style="width:150px;" /><br/>
    *验证码:<br/>
  <input name="txt_check" type="text" size=6 maxlength=4 class="input"><img src="checkcode.asp " alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='checkcode.asp?t='+(new Date().getTime());" ><br/>
  <input type="submit" name="Button1" value="确认注册" /></p>
</form>
<img src="images/fanhui.gif" width="16" height="9" /><a href="/">返回首页</a>
<%end if%>

<%if Request("action")="ok" then%>
<%
dim user,password,sjh,qq,urltitle,url,urllx
user=Request("user")
password=Request("password")
sjh=Request("sjh")
qq=Request("qq")
urltitle=Request("urltitle")
url=Request("url")
urllx=Request("urllx")

if user="" then daving=daving&"用户名不可为空<br/>"
if password="" then daving=daving&"密码不可为空<br/>"
if sjh="" then daving=daving&"联系手机不可为空<br/>"
if qq="" then daving=daving&"QQ号不可为空<br/>"
if urltitle="" then daving=daving&"网站名称不可为空<br/>"
if url="http://" then daving=daving&"网站地址不可为空<br/>"
if url="" then daving=daving&"网站地址不可为空<br/>"
if trim(session("validateCode")) <> trim(Request("txt_check")) then 
response.write("验证码错误，请重新输入<br/><a href='reg.asp?uid="&Request("uid")&"'>返回重写</a>")
response.end
end if 
if user="" or password="" or sjh="" or qq="" or urltitle="" or url="" then
Response.write daving
Response.write "<br/><img src='images/fanhui.gif' width='16' height='9' /><a href='reg.asp?uid="&Request("uid")&"'>返回重写</a><br/>"
else

set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='" & user & "' or sjh='"& sjh &"'"
  rs.open sql,conn
if not rs.eof then
if user=rs("username") then
Response.write "该用户名已被注册！<br/>"
end if
if sjh=rs("sjh") then
Response.write "该手机号已被注册！<br/>"
end if 
Response.write "<br/><img src='images/fanhui.gif' width='16' height='9' /><a href='reg.asp?uid="&Request("uid")&"'>返回重写</a><br/>"
else
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username"
  rs.open sql,conn,1,2
rs.addnew
rs("username")=user
rs("password")=HmacMd5(password,2)
rs("sjh")=sjh
rs("qq")=qq
rs("sxuser")=Request("uid")
rs("sxtime")=now()
rs("sid")=HmacMd5(date()&request.cookies("sid")&time,45)
rs.Update

set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from url"
  rs.open sql,conn,1,2
rs.addnew
rs("username")=user
rs("title")=urltitle
url=Replace(url,"http://","")
url=split(trim(url),"/")
rs("url")=trim(url(0))
rs("urllx")=urllx
rs("zt")=session("ik")
rs.Update

Response.write "恭喜您注册成功！<br/><br/><img src='images/fanhui.gif' width='16' height='9' /><a href='index.asp'>用户登陆</a><br/><br/>"
end if
end if
%>
<%end if%>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
