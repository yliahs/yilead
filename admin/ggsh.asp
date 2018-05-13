<!--#INCLUDE file="conn.asp"-->
<title>广告审核</title>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--广告审核页面-->
<%
  if Request("action")="" then%>
  <div class="p13">&nbsp;广告审核</div>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("ac")="" then
  sql="select * from ggurl where zt=1 Order By id desc"
  elseif Request("ac")="ysh" then
   sql="select * from ggurl where zt=2 Order By id desc"
   elseif Request("ac")="wtg" then
    sql="select * from ggurl  where zt=3 Order By id desc"
   end if
  rs.open sql,conn,1,2
  if rs.eof then
  %><span class="tab">&nbsp;&nbsp;<a href="ggsh.asp">未审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=ysh">已审核</a><span><a href="ggsh.asp?ac=wtg">未通过</a>&nbsp;&nbsp;</span><a href="ggsh.asp?action=ss">搜索</a></span><%
  Response.write "<p>暂无广告申请数据！</p>"
  %> <p><img src="../images/fanhui.gif"/><a href="ggsh.asp">广告审核</a></p>
  <p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
  Response.end
  end if
   rs.Move((page-1)*ada) 
  dim i
  i=1
%> 
<p><%if Request("ac")="" then%>
<span class="tab"><span>&nbsp;&nbsp;未审核&nbsp;&nbsp;</span><a href="ggsh.asp?ac=ysh">已审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=wtg">未通过</a></span>
<%elseif Request("ac")="ysh" then%>
<span class="tab">&nbsp;&nbsp;<a href="ggsh.asp">未审核</a><span>&nbsp;&nbsp;已审核&nbsp;&nbsp;</span><a href="ggsh.asp?ac=wtg">未通过</a></span>
<%elseif Request("ac")="wtg" then%>
<span class="tab">&nbsp;&nbsp;<a href="ggsh.asp">未审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=ysh">已审核</a><span>未通过&nbsp;&nbsp;</span></span>
<%end if%>
<a href="ggsh.asp?action=ss">&nbsp;搜索</a></p>
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">	
<tr></tr></table>
	<%
	
if not (rs.bof and rs.eof)  then
				Rs.PageSize = 10	'一页N条记录
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
 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">					
<tr>	<%=i%>.广告：<%=rs("ggtitle")%><br/>
    用户：<%=rs("username")%><br/>
    站点：<a href="http://<%=rs("url")%>"><%=rs("urltitle")%></a><br/>
    <a href="ggsh.asp?action=sh&amp;id=<%=rs("id")%>">审核</a>&nbsp;&nbsp;<a href="ggsh.asp?action=wtg&amp;id=<%=rs("id")%>">未通过</a>&nbsp;&nbsp;<a href="ggsh.asp?action=sc&amp;id=<%=rs("id")%>">删除</a>
   
    </tr></table>
<% 
	 	Rs.MoveNext
	  	Next

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
call waphx
%>
 
 <!--广告申请搜索页面-->
 <%if Request("action")="ss" then%>
<div class="p13">&nbsp;广告申请搜索</div> 
<%if Request("ac")="" then%>
  <span class="tab">&nbsp;&nbsp;<a href="ggsh.asp">未审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=ysh">已审核</a><span>未通过&nbsp;&nbsp;</span><a href="ggsh.asp?action=ss">搜索</a></span>
 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
<td width="70%" height="22" align="center">
  <form name="form1" method="post" action="ggsh.asp?action=ss&amp;ac=ok">
 用户名：&nbsp; <input name="user" type="text" size="10" maxlength="30" /><br/>
 渠道编号：
 <input name="qdbh" type="text" size="10" maxlength="20" /><br/>
 </td>
 <td width="30%" height="22" align="center">
<div style="text-align:center; margin-top:2px;"> <span id="Label_Ts"><input type="submit" name="Button1" value="确认搜索" /></span></div>
</td>
</tr>
</table>
  </form> 
 <%else
 if Request("user")="" and Request("qdbh")="" then
 Response.write "<p>搜索信息不能为空!</p>"
 Response.end
 end if
  set rs=Server.CreateObject("ADODB.Recordset")
 if Request("user")<>"" and Request("qdbh")="" then
  sql="select * from ggurl where username like '%"&Request("user")&"%' Order By id desc"
  elseif Request("qdbh")<>"" and Request("user")="" then
   sql="select * from ggurl where qdbh like '%"&Request("qdbh")&"%' Order By id desc"
   elseif Request("user")<>"" and Request("qdbh")<>"" then
    sql="select * from ggurl  where username like '%"&Request("urltitle")&"%' and qdbh like '%"&Request("urltitle")&"%' Order By id desc"
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
 elseif rs("zt")="4" then
 zt="代码已收回"
 end if
 Response.write "<p>"&i&".广告："&rs("ggtitle")&"<br>用户:"&rs("username")&"<br/>站点：<a href='http://"&Replace(rs("url"),"http://","")&"'>"&rs("urltitle")&"</a><br/>状态："&zt&"</p>"
 Response.write "<p><a href='ggsh.asp?action=sh&amp;id="&rs("id")&"'>审核</a>&nbsp;&nbsp;<a href='ggsh.asp?action=wtg&amp;id="&rs("id")&"'>未通过</a>&nbsp;&nbsp;<a href='ggsh.asp?action=sc&amp;id="&rs("id")&"'>删除</a></p>"
 
   rs.movenext
    	 loop
  else
  %> <span class="tab">&nbsp;&nbsp;<a href="ggsh.asp">未审核</a>&nbsp;&nbsp;<a href="ggsh.asp?ac=ysh">已审核</a><span>未通过&nbsp;&nbsp;</span><a href="ggsh.asp?action=ss">搜索</a></span><%
  Response.write "<p>没用搜索到相关数据!</p>"
  end if  
end if%>  
  <p><img src="../images/fanhui.gif"/><a href="ggsh.asp">广告审核</a></p>
 <%end if%>
 
 
 <!-- 广告申请删除页面-->
 <%if Request("action")="sc" then %>
  <div class="p13">&nbsp;申请删除</div>
 <% if Request("ac")="" then
 response.write "<p>您确定要删除该用户的申请吗?<br/><a href='ggsh.asp?id="&Request("id")&"&amp;action=sc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='ggsh.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from ggurl where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>广告申请删除成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='ggsh.asp'>广告审核</a></p>"
 end if
 end if
%> 


<!--广告审核页面-->
<%if Request("action")="sh" then%>
 <div class="p13">&nbsp;广告审核</div>
<% if Request("ac")="" then%>
<%

 set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ggurl  where id="&Request("id")&""
  rs.open sql,conn,1,2 

  %>
  广告名称：<%=rs("ggtitle")%>
 <form name="form1" method="post" action="ggsh.asp?action=sh&amp;ac=ok&amp;id=<%=Request("id")%>">
 <p>连接地址：（没有请留空）</p>
 <input name="ggurl" type="text" size="10" maxlength="30" /><br/>
 <p>渠道编号：（没有请留空）</p>
 <input name="qdbh" type="text" size="10" maxlength="20" /><br/>
 
<p><span id="Label_Ts"><input type="submit" name="Button1" value="确认审核" /></span></p>
<% rs.close
set rs=nothing
%>
<% else
 set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ggurl  where id="&Request("id")&""
  rs.open sql,conn,1,2
rs("ggurl")=Request("ggurl")
rs("qdbh")=Request("qdbh")
rs("zt")=2
rs.Update
rs.close
set rs=nothing
Response.write "<p>审核通过成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='ggsh.asp'>广告审核</a></p>"
end if
%>
<%end if%>
 
 <!--广告审核不通过页面-->
 <%if Request("action")="wtg" then%>
  <div class="p13">&nbsp;广告审核</div>
  <%
 if request("ac")="" then
Response.write "<p>确定审核申请为不通过吗？</p>"
Response.write "<p><a href='ggsh.asp?action=wtg&amp;ac=ok&amp;id="&Request("id")&"'>确定</a>&nbsp;&nbsp;<a href='ggsh.asp'>取消</a></p>" 
 else
  set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from ggurl  where id="&Request("id")&""
  rs.open sql,conn,1,2
rs("zt")=3
rs.Update
rs.close
set rs=nothing
Response.write "<p>审核不通过成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='ggsh.asp'>广告审核</a></p>"
 end if
end if%>
 
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

