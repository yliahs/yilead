<!--#INCLUDE file="conn.asp"-->


<title>财务日志</title>	
</head>
<body>
    <div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->

<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />财务日志</div>
	 

<% 
dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1 
			
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From cwrz where username='"&Request.Cookies("username")&"' Order By time desc"
  rs.open sql,conn,1,2
  
  
  if not rs.eof then
   rs.Move((page-1)*ada) 
  dim i
  i=1
%>
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
					%><p>
					交易金额：<%=rs("money")%><br/>
					交易说明：<%=rs("sm")%><br/>
					交易时间：<%=rs("time")%><br/>
		            -------------</p>
<%
	 	Rs.MoveNext
	  	Next

      %>
	  <%
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='rz.asp?page=" & Page + 1 & "'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='rz.asp?page=" & Page - 1 & "'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
			
			Response.Write("<form action='rz.asp?uid=" & uid & "&amp;run=" & Int((9999) * Rnd() + 1) & "' method='post'>")
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<input name='na' type='button' value='跳页' />")
    		    Response.Write("</form><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
	else
	Response.write "<p class='hongse'>==暂无数据==</p>"
	end if
  
  %>

<p class="px"><img src="images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
