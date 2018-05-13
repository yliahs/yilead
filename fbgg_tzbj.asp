<!--#INCLUDE file="conn.asp"-->
<title>广告管理</title>	
</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>


<%if Request("action")="bj" then%>
 <div class="p13">&nbsp;编辑广告</div>
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  rs.open sql,conn,1,2
if rs.eof then
Response.write "<p>广告不存在！</p>"
Response.end
end if
if Request("ac")="" then %>
 <form name="form1" method="post" action="fbgg_tzbj.asp?action=bj&amp;ac=ok&amp;id=<%=Request("id")%>">
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
  <p>计费类型：
  <select name="gglx">
  <%if rs("gglx")=1 then%>
 <option value="1" selected="selected">点击付费</option>
<option value="2">效果付费</option>
<%else%>
<option value="1">点击付费</option>
<option value="2" selected="selected">效果付费</option>
<%end if%>
  </select></p>
  <p>广告分类：
  <select name="gglb">
   <% set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from gglb Order By px asc"
  rsa.open sql,conn,1,2
   do while ((not rsa.eof))
 if rs("gglb")=rsa("id") then  %>
 <option value="<%=rsa("id")%>" selected="selected"><%=rsa("title")%></option>
 <%else%>
 <option value="<%=rsa("id")%>"><%=rsa("title")%></option> 
 <%end if%>
</p>
  <%rsa.MoveNext
loop 
rsa.close
set rsa=nothing%>  
</select>
 <p>广告语：</p>
 <input name="ggy" type="text" value="<%=rs("ggy")%>"/>
 <p>广告说明：</p>
 <input name="ggsm" type="text" value="<%=rs("ggsm")%>"/>
<p><input type="submit" name="Button1" value="确认编辑" /></form></p>

<% else
title=Request("title")
url=Request("url")
logo=Request("logo")
moneyz=Request("money")
jmoney=Request("jmoney")
gglx=Request("gglx")
ggy=Request("ggy")
ggsm=Request("ggsm")
gglb=Request("gglb")
if title="" then ave=ave&"广告名称不能为空！<br/>"
if url="" then ave=ave&"广告地址不能为空！<br/>"
if moneyz="" then ave=ave&"广告单价不能为空！<br/>"
if round(moneyz,2)<0.07 then ave=ave&"广告单价不能小于等于0.01元!<br/>"
if jmoney="" then ave=ave&"单价说明不能为空！<br/>"
if ggy="" then ave=ave&"广告语不能为空！<br/>"
if ggsm="" then ave=ave&"广告说明不能为空！<br/>" 
if gglx="" then ave=ave&"广告类型不能为空！<br/>"
if gglb="" then ave=ave&"广告分类不能为空！<br/>"
if title="" or url="" or moneyz="" or moneyz<0.01 or jmoney="" or ggy="" or ggsm="" or gglx="" or gglb="" then
Response.write "<p>"&ave&"</p>"
else
rs("title")=title
rs("url")=url
rs("logo")=logo
rs("money")=moneyz
rs("jmoney")=jmoney
rs("gglx")=gglx
rs("ggy")=ggy
rs("ggsm")=ggsm
rs("gglb")=gglb
rs.Update
Response.write "<p>广告编辑成功！</p>"
end if
end if
rs.close
set rs=nothing
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="fbgg.asp?action=gl">广告管理</a></p>
<%end if%>


<%if Request("action")="" then%>
 <div class="p13">&nbsp;流量统计</div>
 <form action="fbgg_tzbj.asp?id=<%=Request("id")%>" method="post">
<p>日期：<input name="time1" type="text" value="<%=date-3%>" size="8"/>至<input name="time2" type="text" value="<%=date%>" size="8"/>
<input name="" type="submit" value="查询" />
</p>
 </form>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  rs.open sql,conn,1,2
if rs.eof then
Response.write "<p>广告不存在！</p>"
Response.end
end if
monaa=rs("money")
rs.close
set rs=nothing
time1=Request("time1")
time2=Request("time2")
if time1="" and time2=""then
time1=date-3
time2=date
end if
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ggfw where ggid='"&Request("id")&"' and time>=#"&time1&"# and time<=#"&time2&"# Order By time desc"
rs.open sql,conn,1,2
if not (rs.eof and rs.bof) then
 page=Request.QueryString("page") 'page值为接受值
if page="" then page=1

rs.pagesize=5 '每页显示的记录数
rs.absolutepage=page '显示当前页等于接收到的页数
 for i= 1 to rs.pagesize 
if rs.eof then
exit for
end if			 
Response.write "<p>广告："&rs("title")&"<br/>点击IP："&rs("ip")&"<br/>点击PV："&rs("pv")&"<br/>支出费用："&rs("ip")&"x"&monaa&"="&rs("ip")*monaa&"元<br/>"
Response.write "日期："&rs("time")&"<br/>--------------</p>"


 rs.movenext
next

 zys=cint(rs.RecordCount/rs.pagesize)
   if cint(zys)>1 then
 if page<2 then
%>上一页&nbsp;&nbsp;<%
else
%><a href="?page=<%=page-1%>&amp;id=<%=Request("id")%>&amp;time1=<%=time1%>&amp;time2=<%=time2%>">上一页</a>&nbsp;&nbsp;<%
end if
if cint(page)<zys then
%><a href="?page=<%=page+1%>&amp;id=<%=Request("id")%>&amp;time1=<%=time1%>&amp;time2=<%=time2%>">下一页</a><%
else
%>下一页<%
end if
end if 
if zys=0 then
zys=1
end if
Response.write "<p>共"&zys&"页/共"&rs.recordCount&"条记录</p>"  

else
Response.write "<p>暂无统计数据！</p>"
end if
rs.close
set rs=nothing
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="fbgg.asp?action=gl">广告管理</a></p>
<%end if%>

<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
