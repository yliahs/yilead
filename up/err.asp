<!--#INCLUDE file="../conn.asp"-->
<title>身份证件上传</title>	
</head>
<body>
    <div id="all">
<%if Request.Cookies("username")="" then%>
<%Response.Redirect "/login.asp"%>
<%else%>
<p><div class="p7" align="center"><img src="/<%=session("img")%>" width="220" height="40" /></div></p>
<p><div class="p31">合作ID:<%=Request.Cookies("id")%>|<a id="Top1_HyperLink1" href="/login.asp?action=zx">注销登陆</a></div></p>
<%end if%>

<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

<div class="p13">&nbsp;<img src="../images/dada.gif" />身份证件上传</div>	
<p>
<%'If request("id")="" Then response.redirect"mobileup.asp"%>
<%if request("id")=1 then %>
   文件名不能为空！<br/>
<%End if%>
<%if request("id")=2 then %>
文件类型错误，上传不成功！<br/>
允许上传的文件类型有:gif,jpg,png,jpeg<br/>
<%End if%>
<%if request("id")=3 then %>
文件过大，上传不成功！<br/>
允许上传的文件大小为:100KB<br/>
<%End if%>
<%if request("id")=4 then %>
证件已经上传！<br/>
<%end if%>
<%if request("id")=5 then %>
   请选择要上传的文件！<br/>
<%End if%>
<%if request("id")=6 then %>
系统错误，请联系管理员！<br/>
错误6，保存上传数据所需文件夹不存在！<br/>
<%End if%>
<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='重新上传'/><a href="index.asp">重新上传</a></p>
----------<br/>
</p>

<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="../db.asp"-->
</div>
</body>
</html>
