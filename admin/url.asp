<!--#INCLUDE file="conn.asp"-->
<title>站点管理</title>
<style type="text/css">
<!--
#dd {
	color: #00F;
}
-->
</style></head>
<body>
<div id="all">   
 
<%if Request.COOkies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=session("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--站点管理首页-->
<%
  if Request("action")="" or Request("action")="wsh" or request("action")="ysh" then%>
  <div class="p13">&nbsp;站点管理</div>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("action")="" then
  sql="select * from url Order By id desc"
  elseif Request("action")="wsh" then
   sql="select * from url where zt=1 Order By id desc"
   elseif Request("action")="ysh" then
    sql="select * from url  where zt=2 Order By id desc"
   end if
  rs.open sql,conn,1,2
 
%> 
<%if Request("action")="" then%>
<p><a href="url.asp?action=ss"><span id="dd">站点搜索</a></span></p>
<span class="tab"> <span>&nbsp;&nbsp;&nbsp;全部&nbsp;&nbsp;</span><a href="url.asp?action=wsh">未审核</a>&nbsp;&nbsp;<a href="url.asp?action=ysh">已审核</a></span>
<%elseif Request("action")="wsh" then%>
<p><a href="url.asp?action=ss"><span id="dd">站点搜索</a></span></p>
<span class="tab">&nbsp;&nbsp;&nbsp;<a href="url.asp">全部</a><span>&nbsp;&nbsp;未审核&nbsp;&nbsp;</span><a href="url.asp?action=ysh">已审核</a></span>
<p><a href="url.asp?action=yjsh">一键自动全部审核</a></p>
<%elseif Request("action")="ysh" then%>
<p><a href="url.asp?action=ss"><span id="dd">站点搜索</a></span></p>
<span class="tab">&nbsp;&nbsp;&nbsp;<a href="url.asp">全部</a>&nbsp;&nbsp;<a href="url.asp?action=wsh">未审核</a><span>已审核&nbsp;&nbsp;</span></span>
<%end if%>
<% if rs.eof then
  Response.Write "<p>暂无数据!</p>"
  response.write "<p><img src='../images/fanhui.gif'/><a href='url.asp'>站点管理</a></p>"
  %><p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
  Response.end
  end if %>
  
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="47%" height="22" align="center">站点名称</td>
	<td width="25%" height="22" align="center">站点类型</td>
	<td width="28%" height="22" align="center">站点审核</td>
	</tr>
	<% 
	rs.Move((page-1)*ada) 
  dim i
  i=1
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
		<td width="47%" height="22" align="center"><a href="http://<%=Replace(rs("url"),"http://","")%>"><%=rs("title")%></a></td>
	<td width="25%" height="22" align="center"><%=rs("urllx")%></td>
    <%if Request("action")="wsh" then%>
    <td width="28%" height="22" align="center"><a href="url.asp?id=<%=rs("id")%>&amp;action=tg" class="STYLE1">通过</a>&nbsp;<a href="url.asp?id=<%=rs("id")%>&action=dh" class="STYLE1">打回</a></td>
    <%elseif Request("action")="ysh" then%>
   <td width="28%" height="22" align="center"><a href="url.asp?id=<%=rs("id")%>&amp;action=dh" class="STYLE1">打回</a>&nbsp;<a href="url.asp?id=<%=rs("id")%>&action=sc" class="STYLE1">删除</a></td>
   <%else%>
    <td width="28%" height="22" align="center"><a href="url.asp?id=<%=rs("id")%>&amp;action=bj" class="STYLE1">编辑</a>&nbsp;<a href="url.asp?id=<%=rs("id")%>&action=sc" class="STYLE1">删除</a></td>
   <%end if%>
	</tr>
    
	
<%
	 	Rs.MoveNext
	  	Next

      %> </table><%
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='url.asp?page=" & Page + 1 & "&amp;action="&Request("action")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='url.asp?page=" & Page - 1 & "&amp;action="&Request("action")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='url.asp?action="&Request("action")&"' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
end if
%>
 
 <!--站点搜索页面-->
 <%if Request("action")="ss" then%>
<div class="p13">&nbsp;搜索站点</div> 
<%if Request("ac")="" then%>
 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
	<td width="70%" height="22" align="center">
  <form name="form1" method="post" action="url.asp?action=ss&amp;ac=ok">
 用户： <input name="user" type="text" size="10" maxlength="20" /><br/>
 网址：
 <input name="url" type="text" size="10" maxlength="20" /><br/>
 站名： <input name="urltitle" type="text" size="10" maxlength="20" /><br/>
 </td>
 <td width="30%" height="22" align="center">
<div style="text-align:center; margin-top:2px;"> <span id="Label_Ts"><input type="submit" name="Button1" value="确认搜索" /></span></div>
</td>
</tr>
</table>
  </form> 
 <%else
 if Request("user")="" and Request("url")="" and Request("urltitle")="" then
 Response.write "<p>信息不能为空!</p>"
 Response.end
 end if
  set rs=Server.CreateObject("ADODB.Recordset")
 if Request("user")<>"" then
  sql="select * from url where username like '%"&Request("user")&"%' Order By id desc"
  elseif Request("url")<>"" then
   sql="select * from url where url like '%"&Request("url")&"%' Order By id desc"
   elseif Request("urltitle")<>"" then
    sql="select * from url  where title like '%"&Request("urltitle")&"%' Order By id desc"
   end if
  rs.open sql,conn,1,2
  if not rs.bof then
  i=0
 do while not rs.eof
 i=i+1
 dim zt
 if rs("zt")="1" then
 zt="未审核"
 elseif rs("zt")="2" then
 zt="审核通过"
 elseif rs("zt")="3" then
 zt="审核未通过"
 end if
 Response.write "<p>"&i&".站名："&rs("title")&"<br>网址：<a href='http://"&Replace(rs("url"),"http://","")&"'>"&rs("url")&"</a><br/>用户:"&rs("username")&"<br/>状态："&zt&"</p>"
 Response.write "<p><a href='url.asp?action=bj&amp;id="&rs("id")&"'>编辑</a>&nbsp;&nbsp;<a href='url.asp?action=sc&amp;id="&rs("id")&"'>删除</a>&nbsp;&nbsp;<a href='url.asp?action=tg&amp;id="&rs("id")&"'>通过</a>&nbsp;&nbsp;<a href='url.asp?action=dh&amp;id="&rs("id")&"'>打回</a></p>------------<br/>"
   rs.movenext
    	 loop
  else
  Response.write "<p>没用搜索到相关数据!</p>"
  end if  
end if%>  
  <p><img src="../images/fanhui.gif"/><a href="url.asp">站点管理</a></p>
 <%end if%>
 
 <!--站点编辑页面-->
 <%if Request("action")="bj" then%>
 <div class="p13">&nbsp;编辑站点</div> 
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from url where id="&Request("id")&""
  rs.open sql,conn,1,2
  if Request("ac")="" then
  %>
  <form action="url.asp?action=bj&amp;ac=ok&amp;id=<%=request("id")%>" method="post">
  <p>站名:</p>
  <input name="title" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("title")%>" size="12" />
   <p>网址:</p>
  <input name="url" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("url")%>" size="12" />
   <p>类型:</p>
  <input name="urllx" type="text" id="TextBox1" style="border-color:Yellow;width:80%;" value="<%=rs("urllx")%>" size="12" />
  <p>状态:
  <select name="zt" size="1">
  <%if rs("zt")="1" then%>
    <option value="1" selected="selected">未审核</option>
    <option value="2">已通过</option>
	<option value="3">未通过</option>
    <%elseif rs("zt")="2" then%>
      <option value="1">未审核</option>
    <option value="2" selected="selected">已通过</option>
	<option value="3">未通过</option>
	<%elseif rs("zt")="3" then%>
	 <option value="1">未审核</option>
    <option value="2">已通过</option>
	<option value="3" selected="selected">未通过</option>
    <%end if%>
  </select></p>
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Button1" value="确认修改" id="Button1"/><br/>
  </form>
  <%else
  rs("title")=Request("title")
  rs("url")=Request("url")
  rs("urllx")=request("urllx")
  rs("zt")=request("zt")
  rs.Update
  rs.close
  set rs=nothing
  Response.write "<p>站点修改成功！</p>"
 end if%>
  <p><img src="../images/fanhui.gif"/><a href="url.asp">站点管理</a></p>
  <%  end if%>
 
 

 
 <!-- 站点删除页面-->
 <%if Request("action")="sc" then %>
  <div class="p13">&nbsp;，站点删除</div>
 <% if Request("ac")="" then
 response.write "<p>您确定要删除此站点吗?<br/><a href='url.asp?id="&Request("id")&"&amp;action=sc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='user.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from url where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>站点删除成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='url.asp'>站点管理</a></p>"
 end if
 end if
 call waphx
%> 


<!--站点审核页面-->
<%if Request("action")="tg" then%>
 <div class="p13">&nbsp;站点审核</div>
<% if Request("ac")="" then



Response.write "<p>确定审核通过此站点吗？</p>"
Response.write "<p><a href='url.asp?action=tg&amp;ac=ok&amp;id="&Request("id")&"'>确定</a>&nbsp;&nbsp;<a href='url.asp'>取消</a></p>"
else
 set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from url  where id="&Request("id")&""
  rs.open sql,conn,1,2
rs("zt")=2
rs.Update
Response.write "<p>审核通过成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='url.asp'>站点管理</a></p>"
end if
%>
<%end if%>
 
 <!--站点审核不通过页面-->
 <%if Request("action")="dh" then%>
  <div class="p13">&nbsp;站点审核</div>
  <%
 if request("ac")="" then
Response.write "<p>确定打回此站点吗？</p>"
Response.write "<p><a href='url.asp?action=dh&amp;ac=ok&amp;id="&Request("id")&"'>确定</a>&nbsp;&nbsp;<a href='url.asp'>取消</a></p>" 
 else
  set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from url  where id="&Request("id")&""
  rs.open sql,conn,1,2
rs("zt")=3
rs.Update
Response.write "<p>审核未通过成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='url.asp'>站点管理</a></p>"
 end if
end if%>

<%if Request("action")="yjsh" then%>
<div class="p13">&nbsp;一键自动审核</div>
<%
set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from url where zt=1 Order By id desc"
  rs.open sql,conn,1,2
do while not rs.eof
rs("zt")=2
rs.update
rs.movenext 
  loop 
rs.close
set rs=nothing
Response.write "<p>全部站点审核通过成功！</p>"

%>
<%end if%>
 
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

