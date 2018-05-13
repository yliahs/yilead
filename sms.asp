<!--#INCLUDE file="conn.asp"-->
<title>短消息</title>	

<style type="text/css">
<!--
#ddd {
	color: #00F;	
	}
-->
</style>
</head>
<body>
<div id="all">       
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>
<%if Request("action")="" then%>
      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />短消息</div>
  <div></div>
  
  <%
  if Request("ac")="" then%>
  <span class="tab"> <span>&nbsp;&nbsp;&nbsp;收件箱&nbsp;&nbsp;</span><a href="sms.asp?ac=fjx">发件箱</a></span>
  <%end if%>
   <%if Request("ac")="fjx" then%>
  <span class="tab"> &nbsp;&nbsp;&nbsp;<a href="sms.asp">收件箱</a>&nbsp;&nbsp;<span>发件箱</span></span>
  <%end if%>
  <%
   set rs=Server.CreateObject("ADODB.Recordset")
   if Request("ac")="" then
  sql="SELECT * From sms where sxuser='"&Request.Cookies("username")&"'"
  else
  sql="SELECT * From sms where username='"&Request.Cookies("username")&"'"
  end if
  rs.open sql,conn
 %>
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
    <tr>
    <%if Request("ac")="" then%>
      <td width="20%" height="22" align="center">发信人</td>
      <%else%>
      <td width="20%" height="22" align="center">收信人</td>
      <%end if%>
    <td width="18%" height="22" align="center">属性</td>
	<td width="62%" height="22" align="center">消息标题</td>
    </tr>
	<%
	if not rs.eof then
	rs.MoveFirst
    While Not rs.EOF
	dim zt
	if Request("ac")="" then
	if rs("zt")="0" then
	zt="未读"
	else
	zt="以读"
	end if
	else
	if rs("hh")="1" then
	zt="未复"
	else
	zt="以复"
	end if
	end if
	%>
    <tr>
     <%if Request("ac")="" then%>
      <td width="20%" height="22" align="center"><%=rs("username")%></td>
      <%else%>
       <td width="20%" height="22" align="center">管理员</td>
      <%end if%>
    <td width="18%" height="22" align="center"><%=zt%></td>
    <%if Request("ac")="" then%>
	<%if len(rs("title"))>9 then %>
	<td width="62%" height="22" align="center"><a href="sms.asp?action=ad&amp;id=<%=rs("id")%>"><span class="STYLE2"><%=left(rs("title"),9)%>...</span></a></td>
	<%else%>
	<td width="60%" height="22" align="center"><a href="sms.asp?action=ad&amp;id=<%=rs("id")%>"><span class="STYLE2"><%=rs("title")%></span></a></td>
	<%end if%>
    <%else%>
   <td width="62%" height="22" align="center"><a href="sms.asp?action=ad&amp;id=<%=rs("id")%>&amp;ac=ok"><span class="STYLE2"><%=left(rs("title"),9)%>...</span></a></td> 
    <%end if%>
    </tr>
	<%rs.MoveNext
      Wend
      rs.close
      set rs=nothing
	  else
	  %>
	  <tr>
<td height="22" colspan="3" align="center">
您没有短消息！</td>
	</tr>
	<%end if%>
  </table>
  <%if Request("ac")="fjx" then%>
<p><a href="sms.asp?action=fjx">&nbsp;&nbsp; <span id="ddd">我要发短消息</a></span></p>
  <%end if%>
 <%end if%>
 
 
   <!--发送短消息页面-->
 <%if Request("action")="fjx" then %>
 <div class="p12">发送短消息</div>
<% if Request("ac")="" then %>
  
      <p>收信人：管理员</p>
     <form name="form1" method="post" action="sms.asp?action=fjx&amp;ac=ok">
   <p> 收件人用户名：(可填写"管理员")</p>
       <input name="name" type="text" maxlength="30" />   
   <p> 短信标题：</p>
       <input name="title" type="text" maxlength="30" />
   <p> 短信内容：</p>
     <input name="xxnl" type="text" /><br/>
<span id="Label_Ts"><input type="submit" name="Button1" value="确认发送" /></span>
  </form> 
  <%else
  if request("name")="" or  Request("title")="" or Request("xxnl")="" then
  Response.write "<p>收件人或短信标题或短信内容不能为空!</p>"
  else
  
if request("name")=Request.Cookies("username") then
response.write "不可以给自己发信"
response.end
end if
  sxz=request("name")
if sxz<>"管理员" then
 set rst=Server.CreateObject("ADODB.Recordset")
 rst.open"select * from username where username='"&request("name")&"'",Conn,1,1
 if rst.eof then
 response.write "不存在这个人！"
 response.end
 else
 sxz=request("name")
 end if
 rst.close
 set rst=nothing
 end if
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From sms"
  rs.open sql,conn,1,2
  rs.addnew
  rs("sxuser")=sxz
  rs("username")=Request.Cookies("username")
  rs("title")=Request("title")
  rs("xxnl")=Request("xxnl")
  rs("hh")=1
  rs.update
  rs.close
  set rs=nothing
  Response.write "<p>短消信发送成功！请等待回复！</p>"
  end if
  end if%>
 <%end if%>
 
 
 <!--查看消息页面-->
 <%if Request("action")="ad" then%>
  <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />查看短消息</div>
 <%
  set rs=Server.CreateObject("ADODB.Recordset")
  if Request("ac")="" then
  sql="SELECT * From sms where sxuser='"&Request.Cookies("username")&"' and id="&Request("id")&""
  else
  sql="SELECT * From sms where username='"&Request.Cookies("username")&"' and id="&Request("id")&""
  end if
  rs.open sql,conn,1,2
  if Request("ac")="" then
  rs("zt")=1
  rs.update
   end if
 %>
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
 <tr>
<td width="20%" height="22" align="center">标题</td>
<td width="80%" height="22" align="center"><%=rs("title")%></td>
	</tr>
	 <tr>
<td width="20%" height="22" align="center">内容</td>
<td width="80%" height="22" align="center"><%=rs("xxnl")%></td>
	</tr>
    <%if Request("ac")<>"" then%>
    <tr> <td height="22" colspan="2" align="center">回复：
    <%if rs("hh")="1" then%>
    未复
    <%else%>
   <%=rs("hh")%>
    <%end if%>
    </td></tr>
    <%end if%>
	</table>
<br/><div class="px">&nbsp;<img src="images/fanhui.gif" width="16" height="9"/><a href="sms.asp">消息中心</a></div>
 <%end if%>
 <p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
