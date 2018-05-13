<!--#INCLUDE file="conn.asp"-->
<title>点击IP查询</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
 <div class="p13">&nbsp;点击IP查询</div> 
  
 <p>
 <%
 if time1="" then
 time1=date()-1
 else
 time1=time1
 end if%>
  <form name="form1" method="post" action="djip.Asp?action=cx">
日期： <input name="time1" type="text" id="TextBox1" style="width:120px;" value="<%=time1%>" size="5" maxlength="20" /><br/>
<input name="dfdf" type="submit" id="dfdf" value="确定查询" />
  </form> </p>
 
 <%
 Response.Write "<p>日期："&Request("time1")&"</p>" 
  dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 time1=Request("time1")
 if time1="" then
 time1=date-4
 end if
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ip where time=#"&time1&"# Order By id desc"
  rs.open sql,conn,1,2
  'i=0
'do while (not rs.EOF)
 'i=i+1
 if not (rs.bof and rs.eof)  then
 rs.Move((page-1)*ada) 
  dim i
  i=1
  
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
		
 
set rsz=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ad where id="&rs("ggid")&""
  rsz.open sql,conn
 if not rsz.eof then
 ggmc=rsz("title")
 else
 ggmc="未知！广告已删除"
 end if
 Response.write ""&i&".IP："&rs("ip")&"<br/>点击用户："&rs("username")&"<br/>"
  if not rsz.eof then
  Response.write "被点广告："&ggmc&"<br/>"
  end if
 rsz.close
set rsz=nothing
 if rs("zt")="0" then
 Response.write "状态：未激活<br/>"
 else
 Response.write "状态：已激活<br/>"
 end if
 Response.write "来源地址："&rs("url")&"<br/>"
 Response.write "------------<br/>"
 'rs.MoveNext
'loop

	Rs.MoveNext
	  	Next

    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='djip.asp?page=" & Page + 1 & "&amp;action="&Request("action")&"&amp;time1="&Request("time1")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='djip.asp?page=" & Page - 1 & "&amp;action="&Request("action")&"&amp;time1="&Request("time1")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
			Response.Write("<form action='djip.asp?action="&Request("action")&"&amp;time1="&Request("time1")&"' method='post'>")
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<input name='tj' type='submit' id='ddd' value='跳页'/>")
    		    Response.Write("</form>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		
		Rs.close
	set rs=nothing
else
Response.write "<p>==暂无数据==</p>"
end if
%>
 
 <p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
