<!--#INCLUDE file="conn.asp"-->
<title>广告管理</title>	
</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

  	  
<% if Request("action")="" then%>
  <div class="p13">&nbsp;发布广告</div>
  <%if Request("ac")="" then%>
   <form name="form1" method="post" action="fbgg.asp?ac=ok">
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
<p><input type="submit" name="Button1" value="确认发布" /></form></p>
<%end if%>
<% if Request("ac")="ok" then

' dim title,url,logo,money,jmoney,gglx,ggy,ggsm,mee,tz,gglb
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
if round(moneyz,2)<0.01 then ave=ave&"广告单价不能小于等于0.01元!<br/>"
if jmoney="" then ave=ave&"单价说明不能为空！<br/>"
if ggy="" then ave=ave&"广告语不能为空！<br/>"
if ggsm="" then ave=ave&"广告说明不能为空！<br/>" 
if gglx="" then ave=ave&"广告类型不能为空！<br/>"
if gglb="" then ave=ave&"广告分类不能为空！<br/>"
if title="" or url="" or moneyz="" or moneyz<0.01 or jmoney="" or ggy="" or ggsm="" or gglx="" or gglb="" then
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
rs("username")=Request.Cookies("username")
rs("ggzt")=2
rs("ggsh")=session("ggo")
rs("gglb")=gglb
rs.Update
rs.close
set rs=nothing
Response.write "<p>广告发布成功!请到广告管理激活推广广告。</p>"
end if 
end if%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="fbgg.asp?action=gl">广告管理</a></p>
<%end if%>
 	  
	
<%if Request("action")="gl" then%>
<div class="p13">&nbsp;广告管理</div>
<p>您发布的广告如下：</p>
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' Order By id asc"
  rs.open sql,conn,1,2
  if not (rs.eof and rs.bof) then
  do while ((not rs.eof))
 Response.write "<p>广告名称："&rs("title")&"<br/>广告单价："&rs("money")&"/"&rs("jmoney")&"<br/>广告余额："&rs("usermoney")&"元<br/>"
 if rs("ggzt")=2 then
Response.write "<a href='fbgg.asp?action=jh&amp;id="&rs("id")&"'>激活推广</a>" 
else
Response.write "<a href='fbgg.asp?action=zt&amp;id="&rs("id")&"'>暂停推广</a>"
end if
Response.write "&nbsp;<a href='fbgg_tzbj.asp?id="&rs("id")&"'>流量统计</a>&nbsp;"
Response.write "<a href='fbgg_tzbj.asp?action=bj&amp;id="&rs("id")&"'>编辑</a>&nbsp;<a href='fbgg.asp?action=sc&amp;id="&rs("id")&"'>删除</a><br/>"
Response.write "--------------</p>"
  rs.MoveNext
loop 
else
Response.write "<p>您没有发布的广告！</p>"
end if
rs.close
set rs=nothing
%>
<%end if%>	

<%if Request("action")="jh" then%>
<div class="p13">&nbsp;激活广告</div>
<%set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  rs.open sql,conn,1,2
  if rs.eof and rs.bof then
  Response.write "<p>广告不存在！</p>"
  Response.end
  end if%>
 <%if rs("usermoney")=<0 then%>
<%if Request("ac")="" then%>
<p>侍激活广告：<%=rs("title")%></p>
<form action="fbgg.asp?action=jh&amp;ac=ok&amp;id=<%=Request("id")%>" method="post">
<p>总投放费用：<input name="moneyy" type="text" size="6" />
元<br/>
(最低1元起！)<br/>
<input name="" type="submit" value="确定激活" />
<p>
</form>
<%else
set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&Request.Cookies("username")&"'"
  rsa.open sql,conn,1,2
  if not rsa.eof then
  moneya=rsa("money")
  end if
  if cint(Request("moneyy"))<1 then
  Response.write "<p>最低投入费用不能少于1元</p>"
  else
 if int(moneya)<int(Request("moneyy")) then
 Response.write "<p>您当前的金额不足"&Request("moneyy")&"元!</p>"
 else
 rsa("money")=rsa("money")-Request("moneyy")
 rsa.update
 rsa.close
 set rsa=nothing
 rs("usermoney")=Request("moneyy")
 rs("ggzt")=1
 rs("ggsh")=session("ggo")
 rs.update
 rs.close
 set rs=nothing
Response.write "<p>广告激活成功，开始投放推广中！</p>" 
 end if 
 end if
end if
else
rs("ggzt")=1
 rs.update
 rs.close
 set rs=nothing
Response.write "<p>广告激活成功，开始投放推广中！</p>" 
end if
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="fbgg.asp?action=gl">广告管理</a></p>
<%end if%>	  
	  
<%if Request("action")="zt" then%>
<div class="p13">&nbsp;暂停推广</div>
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  rs.open sql,conn,1,2
  if Request("ac")="" then
  if not rs.eof then
 Response.write "<p>确定暂停推广"""&rs("title")&"""广告吗？<br/><a href='fbgg.asp?action=zt&amp;ac=ok&amp;id="&Request("id")&"'>确定暂停推广</a></p>" 
  else
 Response.write "<p>广告不存在！</p>" 
 Response.end
  end if
  else
 rs("ggzt")=2
  rs.update
  Response.write "<p>广告暂停推广成功！</p>"
  end if
  rs.close
  set rs=nothing
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="fbgg.asp?action=gl">广告管理</a></p>
<%end if%>	


<%if Request("action")="sc" then%>
<div class="p13">&nbsp;删除广告</div>
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  rs.open sql,conn,1,2
  if Request("ac")="" then
  if not rs.eof then
 Response.write "<p>确定删除"""&rs("title")&"""广告吗？<br/><a href='fbgg.asp?action=sc&amp;ac=ok&amp;id="&Request("id")&"'>确定删除广告</a></p>" 
  else
 Response.write "<p>广告不存在！</p>" 
 Response.end
  end if
  else
rs.delete
  Response.write "<p>广告删除成功！</p>"
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
