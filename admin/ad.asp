<!--#INCLUDE file="conn.asp"-->
<title>添加广告</title>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--广告添加页面-->
<%
  if Request("action")="" then%>
  <div class="p13">&nbsp;添加广告</div>
  <%if Request("ac")="" then%>
   <form name="form1" method="post" action="ad.asp?ac=ok">
 <p>广告名称：</p> 
 <input name="title" type="text" maxlength="25" />
 <p>广告地址：</p>
 <input name="url" type="text" value="http://"/>
<p>广告图片：</p>
 <input name="logo" type="text" />
<p>广告单价：(元)</p>
 <input name="money" type="text" />
 <p>单价说明（如激活,安装,日独立,有效IP）</p>
  <input name="jmoney" type="text" />
  <p>计费类型：
  <select name="gglx">
 <option value="1" selected="selected">点击付费</option>
<option value="2">效果付费</option>
  </select></p>
  <p>广告分类：
  <select name="gglb">
   <% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb Order By px asc"
  rs.open sql,conn,1,2
   do while ((not rs.eof))%>
  
 <option value="<%=rs("id")%>"><%=rs("title")%></option>
</p>
  <%rs.MoveNext
loop 
rs.close
set rs=nothing%>  
</select>
 <p>广告语：</p>
 <input name="ggy" type="text" />
 <p>广告说明：</p>
 <input name="ggsm" type="text" />
  <p>广告审核：<select name="ggsh">
 <option value="1" selected="selected">自动审核</option>
<option value="2">手动审核</option>
  </select></p>
  <p>广告状态：<select name="ggzt">
 <option value="1" selected="selected">正常推广</option>
<option value="2">暂停推广</option>
  </select></p>
<p><input type="submit" name="Button1" value="确认添加" /></form></p>
<%end if%>
<% if Request("ac")="ok" then

' dim title,url,logo,money,jmoney,gglx,ggy,ggsm,gglb
title=Request("title")
url=Request("url")
logo=Request("logo")
moneyz=Request("money")
jmoney=Request("jmoney")
gglx=Request("gglx")
ggy=Request("ggy")
ggsm=Request("ggsm")
gglb=Request("gglb")
ggzt=Request("ggzt")
ggsh=Request("ggsh")
if title="" then ave=ave&"广告名称不能为空！<br/>"
if url="" then ave=ave&"广告地址不能为空！<br/>"
if moneyz="" then ave=ave&"广告单价不能为空！<br/>"
if moneyz<0 then ave=ave&"广告单价不能小于等于0!"
if jmoney="" then ave=ave&"单价说明不能为空！<br/>"
if ggy="" then ave=ave&"广告语不能为空！<br/>"
if ggsm="" then ave=ave&"广告说明不能为空！<br/>" 
if gglx="" then ave=ave&"广告类型不能为空！<br/>"
if gglb="" then ave=ave&"广告分类不能为空！<br/>"
if title="" or url="" or moneyz="" or jmoney="" or ggy="" or ggsm="" or gglx="" or gglb="" then
Response.write "<p>"&ave&"</p>"
else
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ad"
  rs.open sql,conn,1,2
rs.addnew
rs("title")=title
rs("url")=url
rs("logo")=logo
rs("money")=moneyz
rs("jmoney")=jmoney
rs("gglx")=gglx
rs("ggy")=ggy
rs("ggsm")=ggsm
rs("gglb")=gglb
rs("ggsh")=ggsh
rs("ggzt")=ggzt
rs.Update
rs.close
set rs=nothing
Response.write "<p>广告添加成功!</p>"
end if 
end if%>
<%end if%>
 
 <!--广告列表页面-->
 <%if Request("action")="ggbj" then%>
<div class="p13">&nbsp;广告列表</div> 
<%  set rs=Server.CreateObject("ADODB.Recordset")
if Request("bbe")="" then%>
<p><span class="tab"><span>正常推广</span>&nbsp;&nbsp;<a href="ad.asp?action=ggbj&bbe=zt">暂停推广</a></span><br/>--------------</p>
<% sql="select * from ad where ggzt=1 Order By id desc"
else%>
<p><span class="tab"><a href="ad.asp?action=ggbj">正常推广</a>&nbsp;&nbsp;<span>暂停推广</span></span><br/>--------------</p>
<% sql="select * from ad where ggzt=2 Order By id desc"
end if%>
<%  
  rs.open sql,conn,1,2
  if not rs.eof then
   i=0
 do while not rs.eof
 i=i+1
 Response.write "<p>"&i&"."&rs("title")&"</p>"
 Response.write "<p><a href='ad.asp?action=ggxg&amp;id="&rs("id")&"'>编辑</a>&nbsp;&nbsp;<a href='ad.asp?action=ggsc&amp;id="&rs("id")&"'>删除</a></p>"
  rs.movenext
    	 loop
 rs.close
 set rs=nothing
 else
 Response.write "<p>暂无广告！</P>"
 end if
end if%>
 

<!--广告修改页面-->
<%if Request("action")="ggxg" then%>
<div class="p13">&nbsp;广告编辑</div> 

<%
 set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ad where id="&Request("id")&""
  rs.open sql,conn,1,2
  
if Request("ac")="" then
  if rs.eof then
  Response.write "<p>无此广告！</p>"
  Response.end
  end if
%>
  <form name="form1" method="post" action="ad.asp?action=ggxg&amp;ac=ok&amp;id=<%=Request("id")%>">
 <p>广告名称：</p> 
 <input name="title" type="text" value="<%=rs("title")%>" maxlength="25" />
 <p>广告地址：</p>
 <input name="url" type="text" value="<%=rs("url")%>"/>
<p>广告图片：</p>
 <input name="logo" type="text" value="<%=rs("logo")%>" />
<p>广告单价：(元)</p>
 <input name="money" type="text" value="<%=rs("money")%>" />
 <p>单价说明（如激活,安装,日独立,有效IP）</p>
  <input name="jmoney" type="text" value="<%=rs("jmoney")%>" />
  <p>广告类型：
  <select name="gglx">
  <%if rs("gglx")="1" then%>
 <option value="1" selected="selected">点击广告</option>
<option value="2">效果广告</option>
<%else%>
 <option value="1">点击广告</option>
<option value="2" selected="selected">效果广告</option>
<%end if%>
  </select></p>
<p>广告分类：
  <select name="gglb">
   <% set rsz=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb Order By px asc"
  rsz.open sql,conn,1,2
   do while ((not rsz.eof))%>
  <% if cint(rs("gglb"))=cint(rsz("id")) then %>
 <option value="<%=rsz("id")%>" selected="selected"><%=rsz("title")%></option>
 <%else%>
 <option value="<%=rsz("id")%>"><%=rsz("title")%></option>
 <%end if%>
</p>
  <%rsz.MoveNext
loop 
rsz.close
set rsz=nothing%>  
</select>
 <p>广告语：</p>
 <input name="ggy" type="text" value="<%=rs("ggy")%>" />
 <p>广告说明：</p>
 <input name="ggsm" type="text" value="<%=rs("ggsm")%>" />
   <p>广告审核：<select name="ggsh">
   <%if rs("ggsh")=1 then%>
 <option value="1" selected="selected">自动审核</option>
<option value="2">手动审核</option>
<%else%>
<option value="1">自动审核</option>
<option value="2" selected="selected">手动审核</option>
<%end if%>
  </select></p>
  <p>广告状态：<select name="ggzt">
  <%if rs("ggzt")=1 then%>
 <option value="1" selected="selected">正常推广</option>
<option value="2">暂停推广</option>
<%else%>
 <option value="1">正常推广</option>
<option value="2" selected="selected">暂停推广</option>
<%end if%>
  </select></p>
<p><input type="submit" name="Button1" value="确认编辑" /></form></p>
<%else
title=Request("title")
url=Request("url")
logo=Request("logo")
moneyz=Request("money")
jmoney=Request("jmoney")
gglx=Request("gglx")
ggy=Request("ggy")
ggsm=Request("ggsm")
gglb=Request("gglb")
ggzt=Request("ggzt")
ggsh=Request("ggsh")
if title="" then ave=ave&"广告名称不能为空！<br/>"
if url="" then ave=ave&"广告地址不能为空！<br/>"
if moneyz="" then ave=ave&"广告单价不能为空！<br/>"
if moneyz<0 then ave=ave&"广告单价不能小于等于0!"
if jmoney="" then ave=ave&"单价说明不能为空！<br/>"
if ggy="" then ave=ave&"广告语不能为空！<br/>"
if ggsm="" then ave=ave&"广告说明不能为空！<br/>" 
if gglx="" then ave=ave&"广告类型不能为空！<br/>"
if gglb="" then ave=ave&"广告分类不能为空！<br/>"

if title="" or url="" or moneyz="" or jmoney="" or ggy="" or ggsm="" or gglx="" or gglb="" then
Response.write "<p>"&ave&"</p>"
else
rs("title")=title
rs("url")=url
rs("logo")=logo
rs("money")=moneyz
rs("jmoney")=jmoney
rs("gglx")=gglx
rs("gglb")=gglb
rs("ggy")=ggy
rs("ggsm")=ggsm
rs("ggzt")=ggzt
rs("ggsh")=ggsh
rs.Update
rs.close
set rs=nothing
Response.write "<p>广告修改成功！</p>"
end if
end if%>
<%end if%>
 
 
 <!-- 广告删除页面-->
 <%if Request("action")="ggsc" then %>
  <div class="p13">&nbsp;广告删除</div>
 <% if Request("ac")="" then
 response.write "<p>您确定要删除该广告吗?<br/><a href='ad.asp?id="&Request("id")&"&amp;action=ggsc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='ad.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from ad where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>广告删除成功！</p>"
 end if
 end if
%> 

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

