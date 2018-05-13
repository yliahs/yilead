<!--#INCLUDE file="../conn.asp"-->
<title>上传站长证件</title>
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

<form enctype="multipart/form-data" action ="up.asp" method="post" Accept-charset=utf-8>
<p>支持上传格式：gif,jpg,png,jpeg</p>
<p>文件大小：100KB</p>

<p>用户名：<input type="text" name="user" title="文件名称"  value="" size="10">
</p>
<p>文&nbsp;&nbsp; 件：<input name="file1" type="file" title="请选择文件" size="12"></p>
&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit value='上传证件'>
</form>
<br/>注:未知回应请再点击一次[上传]
<br/>

<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="../index.asp">返回首页</a></p>
</div>
</body>
</html>

