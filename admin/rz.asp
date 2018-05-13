<!--#INCLUDE file="conn.asp"-->
<title>财务日志</title>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>

<!--财务日志管理首页-->
<%
  if Request("action")="" then%>
  <div class="p13">&nbsp;财务日志</div>
  <p><a href="rz.asp?action=sc">清空财务日志</a></p>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  

 set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from cwrz Order By id desc"
  rs.open sql,conn,1,2
if rs.eof then
%><p>暂无日志信息!</p>
	  <p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
	Response.End()
	end if %>

	<%
	
	 rs.Move((page-1)*ada) 
  dim i
  i=1
if not (rs.bof and rs.eof)  then
				Rs.PageSize = 6	'一页N条记录
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
<p>用户：<%=rs("username")%><br/>交易金额：<%=rs("money")%><br/>交易说明：<%=rs("sm")%><br/>
交易时间<%=rs("time")%></p>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#33CCFF" class="table" id="ggxq">
<tr></tr></table>
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
%>


<%if Request("action")="sc" then%>
 <div class="p13">&nbsp;财务日志</div>
<%if REquest("ac")="" then
Response.write "<p>确定清空财务日志吗？<br/><a href='rz.asp?action=sc&amp;ac=sc'>确定</a>&nbsp;&nbsp;<a href='rz.asp'>取消</a></p>"
else
set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from cwrz"
  conn.execute sql
  Response.write "<p>财务日志删除成功！</p>"
end if%>
<p><img src="../images/fanhui.gif"/><a href="rz.asp">财务日志</a></p>
<%end if%>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

