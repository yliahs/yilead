<!--#INCLUDE file="conn.asp"-->
<title>证件管理</title>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<div class="p13">&nbsp;证件审核管理</div>

<%action=Request("action")
if action="" then
call inde
elseif action="sh" then
call sh
elseif action="tg" then
call tg
elseif action="wtg" then
call wtg
elseif action="sc" then
call sc
end if
%>

<%Function inde
if Request("ac")="" then
Response.write "<p><a href='sfz.asp?ac=ysh'>已审核</a>&nbsp;未审核&nbsp;<a href='sfz.asp?ac=wtg'>未通过</a>&nbsp;<a href='up/index.asp'>上传证件</a></p>"
elseif Request("ac")="ysh" then
Response.write "<p>已审核&nbsp;<a href='sfz.asp'>未审核</a>&nbsp;<a href='sfz.asp?ac=wtg'>未通过</a>&nbsp;<a href='up/index.asp'>上传证件</a></p>"
elseif Request("ac")="wtg" then
Response.write "<p><a href='sfz.asp?ac=ysh'>已审核</a>&nbsp;<a href='sfz.asp'>未审核</a>&nbsp;未通过&nbsp;<a href='up/index.asp'>上传证件</a></p>"
end if
 dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("ac")="" then
  sql="select * from sfz where zt=1 Order By timee desc"
 elseif Request("ac")="ysh" then
   sql="select * from sfz where zt=2 Order By timee desc"
 elseif Request("ac")="wtg" then
   sql="select * from sfz where zt=3 Order By timee desc"
   end if
  rs.open sql,conn,1,2
  if not rs.eof then
   rs.Move((page-1)*ada) 
  dim i
  i=1
%>

 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>
		<td width="40%" height="22" align="center">用户帐号</td>
	<td width="35%" height="22" align="center">上传时间</td>
	<td width="25%" height="22" align="center">审&nbsp;&nbsp;核</td>
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
		<td width="40%" height="22" align="center"><%=rs("username")%></td>
	<td width="35%" height="22" align="center"><%=rs("timee")%></td>
   <td width="25%" height="22" align="center"><a href="sfz.asp?action=sh&amp;user=<%=rs("username")%>" class="STYLE1">审核</a></td>  
	</tr>
	
<%
	 	Rs.MoveNext
	  	Next

      %> </table><%
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='sfz.asp?page=" & Page + 1 & "'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='sfz.asp?page=" & Page - 1 & "'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
			
			Response.Write("<form action='sfz.asp?uid=" & uid & "&amp;run=" & Int((9999) * Rnd() + 1) & "' method='post'>")
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<input name='na' type='button' value='跳页' />")
    		    Response.Write("</form><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
	else
	Response.write "<p>暂无数据</p>"
	end if
End Function%>


<%Function sh
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from sfz where username='"&Request("user")&"'"
  rs.open sql,conn,1,2
if not rs.eof then
Response.write "<p><img src='../up/pic/"&rs("url")&"' width='150' height='120'/></P>"
if rs("zt")=1 then
zt="未审核"
elseif rs("zt")=2 then
zt="审核通过"
elseif rs("zt")=3 then
zt="审核不通过"
end if
Response.write "<p>证件状态："&zt&"</p>"
Response.write "<p>上传时间："&rs("timee")&"</p>"
Response.write "<p><a href='sfz.asp?user="&Request("user")&"&amp;action=tg'>通过</a>&nbsp;"
Response.write "<a href='sfz.asp?user="&Request("user")&"&amp;action=wtg'>未通过</a>"
Response.write "&nbsp;<a href='sfz.asp?user="&Request("user")&"&amp;action=sc'>删除</a></p>"
else
Response.write "<p>此用户为上传证件！</p>"
end if
rs.close
set rs=nothing
End Function%>

<%Function tg
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from sfz where username='"&Request("user")&"'"
  rs.open sql,conn,1,2
if not rs.eof then
rs("zt")=2
rs.update
rs.close
set rs=nothing
Response.write "<p>审核通过成功！</p>"
else
Response.write "<p>此用户为上传证件！</p>"
end if
End Function%>

<%Function wtg
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from sfz where username='"&Request("user")&"'"
  rs.open sql,conn,1,2
if not rs.eof then
rs("zt")=3
rs.update
rs.close
set rs=nothing
Response.write "<p>审核未通过成功！</p>"
else
Response.write "<p>此用户为上传证件！</p>"
end if
End Function%>


<%Function sc
set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from sfz where username='"&Request("user")&"'"
  conn.execute sql
  Response.write "<p>删除成功！</p>"
End Function%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
