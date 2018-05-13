<!--#INCLUDE file="conn.asp"-->
<title>站长数据</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
 <div class="p13">&nbsp;站长数据询查</div>  
 
 <%action=Request("action")
 if action="" then
 call inde
 elseif action="cx" then
 call cx
 end if
 %>
 
 <%Function inde%>
   <form name="form1" method="post" action="zzsj.Asp?action=cx">
 用户帐号：
 <input name="username" type="text" id="TextBox1" style="width:120px;" value="" size="5" maxlength="20" />
 <br/>
     开始日期：
       <input name="time1" type="text" id="TextBox1" style="width:120px;" value="<%=date()-1%>" size="5" maxlength="20" />
<br/>
     结束日期：
     <input name="time2" type="text" id="TextBox2" style="width:120px;" value="<%=date()-1%>" size="5" maxlength="20" /><br/>
<div style="text-align:center; margin-top:2px;"> <span id="Label_Ts"><input type="submit" name="Button1" value="确认查询" /></span></div>
  </form> 
<%End Function%>

<%Function cx
 dim time1,time2,username
  time1=Request("time1")
  time2=Request("time2")
  username=Request("username")
  if time1="" or time2="" or username="" then
  Response.write "<p>用户名或查询日期不能为空！</p>"
  else 
   set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ggsj where username='"&username&"' and time>=#"&time1&"# and time<=#"&time2&"# Order By time desc"
  rs.open sql,conn
  Response.write "日期:"&time1&"到"&time2&"数据<br/>"
  
  %><table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
     <tr>
    <td width="43%" height="22" align="center">广告名称</td>
   <td width="22%" height="22" align="center">有效数据</td>
    <!--<td width="15%" height="22" align="center">总额</td>-->
    <td width="35%" height="22" align="center">日期</td>
    </tr>
	<%if not rs.eof then%> 
	<%do while ((not rs.EOF))
  set rss=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ad where id="&rs("ggid")&""
  rss.open sql,conn
  if not rss.eof then
  ggmc=rss("title")
  else
  ggmc="未知！广告已删除"
  end if
  %>  
   <tr>
    <td width="30%" height="22" align="center"><%=ggmc%></td>
   <td width="25%" height="22" align="center"><%=rs("zrsj")%></td>
    <td width="30%" height="22" align="center"><%=rs("time")%></td>
    </tr><%rss.close
set rss=nothing
   rs.MoveNext
loop
rs.close
set rs=nothing

  
	else%>
   <tr> <td height="22" colspan="3" align="center">暂无有效数据！</td></tr>
 </table> 
 <%  end if %>
<%  end if
End Function%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p> 
</div>
</body>
</html>
