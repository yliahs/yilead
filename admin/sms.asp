<!--#INCLUDE file="conn.asp"-->
<title>消息管理</title>
<style type="text/css">
<!--
#dd {
	color: #00F;
}
-->
</style></head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--消息管理首页-->
<%
  if Request("action")="" or request("action")="fjx" then%>
  <div class="p13">&nbsp;消息管理</div>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
	dim dd
	dd="管理员"		
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("action")="" then
  sql="select * from sms where sxuser='"&dd&"' Order By id desc"
  else
sql="select * from sms where username='"&Request.Cookies("admin")&"' Order By id desc"  
  end if
  rs.open sql,conn,1,2

%> 
<%if Request("action")="" then%>

<span class="tab"> <span>&nbsp;&nbsp;收件箱&nbsp;&nbsp;</span><a href="sms.asp?action=fjx">发件箱</a></span>
<%elseif Request("action")="fjx" then%>
<span class="tab"> &nbsp;&nbsp;<a href="sms.asp">收件箱</a>&nbsp;&nbsp;<span>发件箱</span></span>
<%end if%>
<% if rs.eof then
	Response.Write "<p>暂无信息!</p>"
	 %>
	  <p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
	Response.End()
	end if %>
	<%
	
	 rs.Move((page-1)*ada) 
  dim i
  i=1
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
	<p>
        <%if Request("action")="" then%>
		<%=i%>.发件人：<%=rs("username")%><br/>
        <%else%>
       	<%=i%>.收信人：<%=rs("sxuser")%><br/>
        <%end if%>
	短信标题：<a href="sms.asp?action=ckxx&amp;id=<%=rs("id")%>"><%=rs("title")%></a><br/>
	短信回复：
	<%if rs("hh")=1 then
	Response.write "未回复"
	else
	Response.write "已回复"
	end if%>
	<br/><a href="sms.asp?action=sc&amp;id=<%=rs("id")%>">删除短信</a><br/>------------</p>
	
<%
	 	Rs.MoveNext
	  	Next

      
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='?page=" & Page + 1 & "&amp;action="&Request("action")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='?page=" & Page - 1 & "&amp;action="&Request("action")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='?action="&Request("action")&"' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
end if
%>
 
 <!--查看消息页面-->
 <%if Request("action")="ckxx" then%>
<div class="p13">&nbsp;查看消息</div> 
<%if Request("ac")="" then

 set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from sms where id="&Request("id")&""  
  rs.open sql,conn,1,2
%>
<p><b> <%=rs("title")%></b></p>------------<br/><p><%=rs("xxnl")%></p>
<p>------------<br/>
  <form name="form1" method="post" action="sms.asp?action=ckxx&amp;ac=ok&amp;id=<%=Request("id")%>">
 回复内容 <input name="hh" type="text" size="10"/><br/>
 <span id="Label_Ts"><input type="submit" name="Button1" value="确认回复" /></span></p>
  </form> 
 <p> 短信回复：
  <%if rs("hh")=1 then%>
  未回复
  <%else
  Response.write rs("hh")
  end if%>
 <%else
 if Request("hh")="" then
 Response.write "<p>信息不能为空!</p>"
 Response.end
 end if
  set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from sms  where id="&Request("id")&""
  rs.open sql,conn,1,2
  if not rs.bof then
 rs("hh")=Request("hh") 
 rs.UPdate
 Response.write "<p>消息回复成功!</p>"
  else
  Response.write "<p>没有此条短信!</p>"
  end if  
end if%>  
  <p><img src="../images/fanhui.gif"/><a href="sms.asp">消息管理</a></p>
 <%end if%>
 
 
 
 <!-- 站点删除页面-->
 <%if Request("action")="sc" then %>
  <div class="p13">&nbsp;，短信删除</div>
 <% if Request("ac")="" then
 response.write "<p>您确定要删除短信吗?<br/><a href='sms.asp?id="&Request("id")&"&amp;action=sc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='sms.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from sms where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>短信删除成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='sms.asp'>消息管理</a></p>"
 end if
 end if
%> 


<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

