<!--#INCLUDE file="conn.asp"-->
<title>后台管理中心</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>

<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<div class="p13">&nbsp;系统信息</div>
易乐广告联盟程序<br/>版本：4.5.13<br/>内核 ：Yile ADwap59<br/>开发：Yliahs<br/>服务端口:<%=Request.ServerVariables("SERVER_PORT")%><br/>网站域名:<%=Request.ServerVariables("HTTP_HOST")%><br/>服务器版本:<%=Request.ServerVariables("SERVER_SOFTWARE")%><br/>服务器IP地址:<%=Request.ServerVariables("LOCAL_ADDR")%>
 <div class="p13">&nbsp;系统管理</div> 
<p> <a href="xtsz.asp">系统设置</a>&nbsp;&nbsp;<a href="urlcs.asp">非法检测</a>&nbsp;&nbsp;<a href="admin.asp">管理帐号</a>&nbsp;&nbsp;<a href="wzgg.asp">网站公告</a><br/>
<a href="sjk.asp">备份还原</a><br/>
</p>
<div class="p13">&nbsp;用户管理</div> 
 <p>
<a href="user.asp">用户管理</a>&nbsp;&nbsp;<a href="url.asp">站点管理</a>&nbsp;&nbsp;<a href="sfz.asp">证件审核</a>&nbsp;&nbsp;<a href="sms.asp">消息管理</a><br/>
</p>
<div class="p13">&nbsp;广告管理</div>
  <p>
  <a href="ggsz.asp">广告设置</a>&nbsp;&nbsp;<a href="ad.asp">广告添加</a>&nbsp;&nbsp;<a href="ad.asp?action=ggbj">广告列表</a>&nbsp;&nbsp;<a href="gglb.asp">广告分类</a><br/>
  <a href="gglb.asp?action=tj">添加分类</a>&nbsp;&nbsp;<a href="ggsh.asp">广告审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=ysh">申请管理</a><br/>
</p>
<div class="p13">&nbsp;帐务管理</div>
  <p>
<a href="adsj.asp">数据录入</a>&nbsp;&nbsp;<a href="ggtj.asp">广告统计</a>&nbsp;&nbsp;<a href="ph.asp">金钱排行</a>&nbsp;&nbsp;<a id="HyperLink2" href="ip.asp">IP清空</a><br/>
<a href="zzsj.asp">站长数据</a>&nbsp;&nbsp;<a id="HyperLink2" href="djip.asp">点击来源</a>&nbsp;&nbsp;<a href="jsgl.asp">结算管理</a>&nbsp;&nbsp;<a href="rz.asp">财务日志</a><br/>
<a href="zhczkc.asp?action=cz">帐户冲值</a>&nbsp;&nbsp;<a href="zhczkc.asp?action=kc">帐户扣除</a><br/>
</p>
</body>
</html>

