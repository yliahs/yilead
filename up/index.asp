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

<%
action=Request("action")
if action="" then
call inde
elseif action="sc" then
call sc
elseif action="xg" then
call sc
end if
%>
<%Function inde
sq1="select * from sfz where username='"&Request.Cookies("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.Recordset")
rs1.open sq1,conn,1,2
if not rs1.eof then
if rs1("zt")=1 then
zt="未审核"
elseif rs1("zt")=2 then
zt="审核通过"
elseif rs1("zt")=3 then
zt="审核不通过"
end if
if rs1("zt")=1 or rs1("zt")=3 then
Response.write "<p><a href='index.asp?action=xg'>修改上传证件</a></p>"
end if 
Response.write "<p>证件状态："&zt&"</p>"
Response.write "<p>上传时间："&rs1("timee")&"</p>"
Response.write "<p>证件浏览：<br/><img src='/up/pic/"&rs1("url")&"' width='150' height='120'/></p>"
else
Response.write "<p>您未上传身份证件！<br/><a href='index.asp?action=sc'>上传身份证件</a></p>"
end if
End Function%>
	 
<%Function sc%>
<form enctype="multipart/form-data" action ="up.asp?action=<%=Request("action")%>" method="post" Accept-Charset="gb2312">
<p>支持上传格式：gif,jpg,png,jpeg</p>
<p>文件大小：100KB</p>
<p>文件:<input name="file1" type="file" title="请选择文件" size="12"></p>
&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit value='上传证件'>
</form>
<br/>注:未知回应请再点击一次[上传]
<br/>
<%End Function%>

<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="../db.asp"-->
</div>
</body>
</html>
