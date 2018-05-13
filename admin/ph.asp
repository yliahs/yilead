<!--#INCLUDE file="conn.asp"-->
<title>站长金钱排行</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<div class="p13">&nbsp;站长金钱排行</div>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username Order By money desc"
  rs.open sql,conn,1,2 
  if rs.eof then
  Response.end
  end if
%> 

<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="60%" height="22" align="center">用户账号</td>
	<td width="40%" height="22" align="center">帐户金额</td>
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

%>
<% 
da=1
		For i=1 to  Rs.PageSize
		If Rs.Eof Then
			exit For
		End If
		da=da+1
					%>
<td width="60%" height="22" align="center"><%=rs("username")%></td>
	 <td width="40%" height="22" align="center"><%=rs("money")%></td>
	</tr>
	<%
	 	Rs.MoveNext
	  	Next

      %>
   </table>
<%
	
    		'分页
    		If Page < rs.PageCount Then
    		Response.Write("<a href='ph.asp?page=" & Page + 1 & "&amp;action="&Request("action")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='ph.asp?page=" & Page - 1 & "&amp;action="&Request("action")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<form name='form1' method='post' action='ph.asp?action="&Request("action")&"'> ")
    		    Response.Write(" <input name='page' type='text' value=''>")
    		    Response.Write("<input type='submit' name='Button1' value='转到' />")
    		    'Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</form?")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing%>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
