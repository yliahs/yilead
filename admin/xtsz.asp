<!--#INCLUDE file="conn.asp"-->
<title>系统设置</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where name='"&Request.Cookies("admin")&"'"
  rs.open sql,conn,1,2
  
  if Request("action")="" then
%>
 <div class="p13">&nbsp;系统设置</div> 
  <form action="xtsz.asp?action=ad" method="post">
  <p>网站名称:</p>
  <input name="title" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("title")%>" size="12" />
  <p>网站地址:</p>
  <input name="url" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("url")%>" size="12" />
   <p>网站LOGO:</p>
  <input name="img" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("img")%>" size="12" />
   <p>网站底部文字:</p>
  <input name="dl" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("dl")%>" size="12" />
   <p>首页文字:</p>
  <input name="diy" type="text" id="TextBox1" style="border-color:Yellow;width:100%;" value="<%=rs("diy")%>" size="14" />
   <p>网站说明:</p>
  <input name="description" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("description")%>" size="12" />
  <p>网站关健字:</p>
  <input name="keywords" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("keywords")%>" size="12" />
   <p>网站滚动公告:</p>
  <input name="gdgg" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("gdgg")%>" size="12" />
   <p>自动审核网站:
  <select name="ik" size="1">
   <%if rs("sfz")=1 then%>
    <option value="1" selected="selected">关闭</option>
    <option value="2">开启</option>
    <%else%>
    <option value="1">关闭</option>
    <option value="2" selected="selected">开启</option>
    <%end if%>
  </select></p>
    <p>自动审核广告:
  <select name="ggo" size="1">
   <%if rs("sfz")=1 then%>
    <option value="2" selected="selected">关闭</option>
    <option value="1">开启</option>
    <%else%>
    <option value="2">关闭</option>
    <option value="1" selected="selected">开启</option>
    <%end if%>
  </select></p>
    <p>身份证件认证:
  <select name="sfz" size="1">
   <%if rs("sfz")=1 then%>
    <option value="1" selected="selected">关闭</option>
    <option value="2">开启</option>
    <%else%>
    <option value="1">关闭</option>
    <option value="2" selected="selected">开启</option>
    <%end if%>
  </select></p>
  <p>网站状态:
  <select name="wzzt" size="1">
  <%if rs("wzzt")="1" then%>
    <option value="1" selected="selected">正常</option>
    <option value="2">关闭</option>
    <%else%>
      <option value="1">正常</option>
    <option value="2" selected="selected">关闭</option>
    <%end if%>
  </select></p>
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Button1" value="确认设置" id="Button1"/><br/>
  </form>
<%
rs.close
set rs=nothing
end if%>


<%if Request("action")="ad" then%>
<div class="p13">&nbsp;系统设置</div>
<% 

rs("title")=Request("title")
rs("url")=Request("url")
rs("img")=Request("img")
rs("dl")=Request("dl")
rs("description")=Request("description")
rs("keywords")=Request("keywords")
rs("sfz")=Request("sfz")
rs("wzzt")=Request("wzzt")
rs("gdgg")=Request("gdgg")
rs("ik")=Request("ik")
rs("ggo")=Request("ggo")
rs.UPdate
rs.close
set rs=nothing
Response.write "<p>系统设置成功！</p>"
end if %>


<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p> 
</div>
</body>
</html>

