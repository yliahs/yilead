<!--#INCLUDE file="../conn.asp"-->
<title>身份证件上传</title>	
</head>
<body>
    <div id="all">
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="../login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<div class="p13">&nbsp;身份证件上传</div>	  
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

<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="../index.asp">返回管理首页</a></p>
</div>
</body>
</html>

