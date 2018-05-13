<!--#INCLUDE file="conn.asp"-->
<title>下线推广</title>	
</head>
<body>
    <div id="all">

<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->

<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />下线推广</div>
	 
<p>你只要把我们提供的推荐地址推荐给其他朋友或站长<br/>成功注册后，该站长广告费结算，你就会有该站长广告费的5%的提成。<br/>如果该站长一直有收入你就一直有提成</p>

<% 
dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1 
			
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From username where sxuser='"&Request.Cookies("id")&"' Order By id desc"
  rs.open sql,conn,1,2
  
  
  if not rs.eof then
   rs.Move((page-1)*ada) 
  dim i
  i=1
%>

 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="50%" height="22" align="center">下线帐号</td>
	<td width="50%" height="22" align="center">注册时间</td>
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
		<td width="50%" height="22" align="center"><%=rs("username")%></td>
	<td width="50%" height="22" align="center"><%=rs("sxtime")%></td>
	</tr>
	
<%
	 	Rs.MoveNext
	  	Next

      %> </table><%
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='xxuser.asp?page=" & Page + 1 & "'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='xxuser.asp?page=" & Page - 1 & "'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
			
			Response.Write("<form action='xxuser.asp?uid=" & uid & "&amp;run=" & Int((9999) * Rnd() + 1) & "' method='post'>")
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<input name='na' type='button' value='跳页' />")
    		    Response.Write("</form><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
	else
	Response.write "<p class='hongse'>==暂无下线会员数据==</p>"
	end if
  
  %>
下线推广地址：<p class="hongse"><%=session("url")%>/reg.asp?uid=<%=Request.Cookies("id")%></p>


<p class="px"><img src="images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
