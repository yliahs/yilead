<!--#INCLUDE file="conn.asp"-->
<title>数据查询</title>	

</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

     <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />数据查询</div>
<%if Request("action")="" then%>
     <div style="text-align:center; margin-top:2px;"> <span id="Label_Ts">选择日期查询</span></div>
     
     <form name="form1" method="post" action="gg_sjcx.Asp?action=ok">
     开始日期：
       <input name="time1" type="text" id="TextBox1" style="width:120px;" value="<%=date()-1%>" size="5" maxlength="20" />
<br/>
     结束日期：
     <input name="time2" type="text" id="TextBox2" style="width:120px;" value="<%=date()-1%>" size="5" maxlength="20" /><br/>
<div style="text-align:center; margin-top:2px;"> <span id="Label_Ts"><input type="submit" name="Button1" value="确认查询" /></span></div>
  </form> 
  <div style=" margin-top:2px;"><br/>今日数据,请在明日12:00的查询!</div> 
  <%end if%>
  
  
  <% if Request("action")="ok" then 
  dim time1,time2
  time1=Request("time1")
  time2=Request("time2")
  if time1="" or time2="" then
  Response.write "<br/>日期不能为空，请返回重新填写!<br/><a href='gg_sjcx.asp'>&nbsp;返回重新输入</a><br/>"
   Response.write "<p class='px'><img src='images/fanhui.gif' width='16' height='9' /><a href='/'>返回首页</a></p>"
%><!--#INCLUDE file="db.asp"--><%
  Response.end
  end if
  if Replace(time2,"-","")>=Replace(date,"-","") then
  time2=date-1
  end if
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ggsj where username='"&Request.Cookies("username")&"' and time>=#"&time1&"# and time<=#"&time2&"# Order By time desc"
  rs.open sql,conn
  Response.write "日期:"&time1&"到"&time2&"数据<br/>"
  %><table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
     <tr>
    <td width="43%" height="22" align="center">广告名称</td>
   <td width="22%" height="22" align="center">有效数据</td>
    <td width="35%" height="22" align="center">日期</td>
    </tr><%
	if rs.eof then
	%><tr> <td height="22" colspan="3" align="center">您暂无有效数据！</td></tr><%
	end if
  do while ((not rs.EOF))
  %>  
   <tr>
    <td width="30%" height="22" align="center"><%=rs("ggtitle")%></td>
   <td width="25%" height="22" align="center"><%=rs("zrsj")%></td>
    <td width="30%" height="22" align="center"><%=rs("time")%></td>
    </tr>    
  <%
   rs.MoveNext
loop
rs.close
set rs=nothing

%></table> 
<br/><p class="px"><img src="images/fanhui.gif" width="16" height="9" /><a href="gg_sjcx.asp">数据查询</a></p> 
<% end if %>
<p class="px"><img src="images/fanhui.gif" width="16" height="9" /><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>

