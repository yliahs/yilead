<!--#INCLUDE file="conn.asp"-->
<title>用户管理</title>
<style type="text/css">
<!--
.STYLE1 {color: #0000FF}
-->
</style>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--用户管理首页-->
<%
  if Request("action")="" then%>
<% dim page
    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("att")="" then
 tte="所有用户列表"
  sql="select * from username Order By id desc"
  elseif Request("att")="zc" then
  tte="正常用户"
  sql="select * from username where zt=0 Order By id desc"
  elseif Request("att")="dj" then
  tte="冻结用户"
  sql="select * from username where zt=1 or zt=2 Order By id desc"
  end if
  rs.open sql,conn,1,2
   rs.Move((page-1)*ada) 
  dim i
  i=1
 %> <div class="p13">&nbsp;<%=tte%></div><%
Response.write "<p><a href='user.asp?att=zc'>正常用户</a>&nbsp;&nbsp;<a href='user.asp?att=dj'>冻结用户</a>&nbsp;&nbsp;<a href='user.asp?action=ss'>用户搜索</a></p>"
%> <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="50%" height="22" align="center">用户帐号</td>
	<td width="25%" height="22" align="center">帐号金额</td>
	<td width="25%" height="22" align="center">编辑删除</td>
	</tr>
	<%
if not (rs.bof and rs.eof)  then
				Rs.PageSize = 15	'一页N条记录
				IF Not IsEmpty(Page) Then
					IF Not IsNumeric(Page) Then		'判断Page是否为数字
						Page=1
					Else
						Page=Cint(Page)		'转换成短整形Integer
					End IF
					IF Page > Rs.PageCount Then
						Rs.AbsolutePage = Rs.PageCount	'设置当前显示页等于最后一页
					ElseIF Page <= 0 Then
						Rs.AbsolutePage = 1		'设置当前页等于第一页
					Else
						Rs.AbsolutePage = Page	'如果大于零,显示当前页等于接收的页数
					End IF
				Else
					Rs.AbsolutePage = 1
				End IF
				Page = Rs.AbsolutePage


		For i=1 to  Rs.PageSize
		If Rs.Eof Then
			exit For
		End If
					%>
					<tr>
		<td width="50%" height="22" align="center"><a href="user.asp?id=<%=rs("id")%>&amp;action=ck"><%=rs("username")%></a></td>
	<td width="25%" height="22" align="center"><%=rs("money")%></td>
   <td width="25%" height="22" align="center"><a href="user.asp?id=<%=rs("id")%>&amp;action=bj" class="STYLE1">编辑</a>&nbsp;<a href="user.asp?id=<%=rs("id")%>&action=sc" class="STYLE1">删除</a></td>  
	</tr>
	
<%
	 	Rs.MoveNext
	  	Next

      %> </table><%
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='user.asp?page=" & Page + 1 & "'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='user.asp?page=" & Page - 1 & "'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='user.asp?uid=" & uid & "&amp;run=" & Int((9999) * Rnd() + 1) & "' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
end if
%>
 
 <!--查看用户信息页面-->
 <%if Request("action")="ck" then%>
<div class="p13">&nbsp;用户信息</div> 
 <%

  set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where id="&Request("id")&""
  rs.open sql,conn
 %><p>ID:<%=rs("id")%><br/>帐号:<%=rs("username")%><br/>密码：<%=rs("password")%><br/>金额：<%=rs("money")%><br/>Q&nbsp;Q：<%=rs("qq")%><br/>邮箱：<%=rs("email")%><br/>手机：<%=rs("sjh")%><br/>开户银行：<%=rs("khh")%><br/>开户地址：<%=rs("khdc")%><br/>开户姓名:<%=rs("khm")%><br/>银行帐号:<%=rs("yhzh")%><br/>帐号状态:
 <%if rs("zt")=0 then
 Response.write "正常"
 elseif rs("zt")=1 then
 Response.write "关闭"
 elseif rs("zt")=2 then
 Response.write "冻结"
 end if
 Response.write "<p><img src='../images/fanhui.gif'/><a href='user.asp'>用户管理</a></p>"
 rs.close
 set rs=nothing
 end if%>

 <!--用户帐号信息修改-->
 <%if Request("action")="bj" then%>
 <div class="p13">&nbsp;用户信息修改</div> 
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where id="&Request("id")&""
  rs.open sql,conn,1,2
  if Request("ac")="" then
  %>
  <form action="user.asp?action=bj&amp;ac=ok&amp;id=<%=request("id")%>" method="post">
  <p>帐号:<%=rs("username")%></p>
  <p>密码:</p>
  <input name="password" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("password")%>" size="12" />
   <p>金额:</p>
  <input name="money" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("money")%>" size="12" />
   <p>Q&nbsp;Q:</p>
  <input name="qq" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("qq")%>" size="12" />
  <p>邮箱:</p>
  <input name="email" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("email")%>" size="12" />
   <p>手机:</p>
 <input name="sjh" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("sjh")%>" size="12" /> 
  <p>开户银行:</p>
 <input name="khh" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("khh")%>" size="12" /> 
  <p>开户地址:</p>
 <input name="khdc" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("khdc")%>" size="12" /> 
  <p>开户姓名:</p>
 <input name="khm" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("khm")%>" size="12" /> 
  <p>银行帐号</p>
 <input name="yhzh" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("yhzh")%>" size="12" /> 
  
  <p>帐号状态:
  <select name="zt" size="1">
  <%if rs("zt")="0" then%>
    <option value="0" selected="selected">正常</option>
    <option value="1">关闭</option>
	<option value="2">冻结</option>
    <%elseif rs("zt")="1" then%>
      <option value="0">正常</option>
    <option value="1" selected="selected">关闭</option>
	<option value="2">冻结</option>
	<%elseif rs("zt")="2" then%>
	 <option value="0">正常</option>
    <option value="1">关闭</option>
	<option value="2" selected="selected">冻结</option>
    <%end if%>
  </select></p>
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Button1" value="确认修改" id="Button1"/><br/>
  </form>
  <p><img src="../images/fanhui.gif"/><a href="user.asp">用户管理</a></p>
  <% 
  end if
  
 if Request("ac")="ok" then
 if rs("password")<>Request("password") and rs("password")<>"" then
 rs("password")=HmacMd5(Request("password"),2)
 end if
 rs("money")=Request("money") 
 rs("qq")=Request("qq")
 rs("email")=Request("email")
 rs("sjh")=Request("sjh")
 rs("khh")=Request("khh")
 rs("khdc")=Request("khdc")
 rs("khm")=Request("khm")
 rs("yhzh")=Request("yhzh")
 rs("zt")=Request("zt")
 rs.update
 rs.close
 set rs=nothing
 Response.write "<p>用户信息修改成功!</p>"
  Response.write "<p><img src='../images/fanhui.gif'/><a href='user.asp'>用户管理</a></p>"
 end if
 end if%>
 
 <!--帐号删除页面-->
 <%if Request("action")="sc" then %>
  <div class="p13">&nbsp;，帐号删除</div>
 <% if Request("ac")="" then
 response.write "<p>您确定要删除此帐号吗?<br/><a href='user.asp?id="&Request("id")&"&amp;action=sc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='user.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from username where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>帐号删除成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='user.asp'>用户管理</a></p>"
 end if
 end if
%> 
 

<!--用户搜索页面-->
<%if request("action")="ss" then%>
 <div class="p13">&nbsp;帐号搜索</div>
<%if Request("ac")="" then%>
请输入帐号：<br/>
<form action="user.asp?action=ss&amp;ac=ok" method="post">
<input name="username" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="" size="20" />
<input type="submit" name="Button1" value="确认搜索" id="Button1"/>
</form>
<%
else
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&Request("username")&"'"
  rs.open sql,conn
if not rs.eof then
Response.write "<p>用户名:"&rs("username")&"<br/>帐号金额："&rs("money")&"<br/>"
Response.write "<a href='user.asp?id="&rs("id")&"&amp;action=bj'>编辑</a>&nbsp;&nbsp;<a href='user.asp?id="&rs("id")&"&amp;action=sc'>删除</a></p>"
else
Response.write "<p>查无数据！</p>"
end if
end if%>
<%end if%>




<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p> 
</div>
</body>
</html>

