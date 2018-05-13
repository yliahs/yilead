<!--#INCLUDE file="conn.asp"-->
<!--#INCLUDE file="../hs.asp"-->
<title>系统设置</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
 <div class="p13">&nbsp;非法站点检测</div>
<%
dim action
action=Request("action")
if action="" then
call inde
elseif action="jc" then
call jc
elseif action="byjc" then
call byjc
elseif action="syjc" then
call syjc
end if
%>

<!--站点管理首页-->
<%Function inde%>
  
 
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

<% if rs.eof then
  Response.Write "<p>暂无数据!</p>"
  response.write "<p><img src='../images/fanhui.gif'/><a href='url.asp'>站点管理</a></p>"
  %><p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
  Response.end
  end if %>
  
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="70%" height="22" align="center">站点名称</td>
	<td width="30%" height="22" align="center">信息检测</td>
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

%><form name="form1" method="post" action="urlcs.asp?action=byjc"> <% 
da=1
		For i=1 to  Rs.PageSize
		If Rs.Eof Then
			exit For
		End If
		da=da+1
					%>
					<tr>
		<td width="70%" height="22" align="center"><a href="http://<%=Replace(rs("url"),"http://","")%>"><%=rs("title")%></a></td>
	 <td width="30%" height="22" align="center"><a href="urlcs.asp?url=<%=rs("url")%>&amp;action=jc" class="STYLE1">检测</a></td>
	</tr>
   
   <input name="url<%=da%>" type="hidden" value="<%="http://"&rs("url")&""%>">
 <input name="title<%=da%>" type="hidden" value="<%=rs("title")%>">
<%
	 	Rs.MoveNext
	  	Next

      %> </table>
	<p> <input type="submit" name="Button1" value="检测本页所有站点" /></p>
</form> <%
	
    		'分页
    		If Page < rs.PageCount Then
    		Response.Write("<a href='urlcs.asp?page=" & Page + 1 & "&amp;action="&Request("action")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='urlcs.asp?page=" & Page - 1 & "&amp;action="&Request("action")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='urlcs.asp?action="&Request("action")&"' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
End Function
%>
 
<%Function jc
url="http://"&Request("url")&""

if xxurl(Url)=False then
Response.write "<p>站点状态：正常访问<p>"
Response.write "<p>此站点无非法关键字信息</p>"
else
Response.write "<p><font color='#ED1F8D'>此站点包含非法关键字<br/>"&xxurl(Url)&"</font></p>"
end if
End Function %>

<%Function byjc
for i=1 to 16

url=Request("url"&i+1)
Response.write "<p>"&Request("title"&i+1)&"</p>"
if  xxurl(url)=False then
Response.write "<p>此站点无非法关键字信息</p>"
else
Response.write "<p><font color='#ED1F8D'>此站点包含非法关键字<br/>"&xxurl(url)&"</font></p>"
end if

next
End Function%>

 
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>