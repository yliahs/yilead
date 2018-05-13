
      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />广告申请</div>
  <div></div>
  <br/>
  <%
  if Request("ac")="" or Request("ac")="cc" then
  Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From url where username='"&Request.Cookies("username")&"' and zt=2 Order By id"
rs.Open sql,conn,1,2
if not(rs.eof and rs.bof) then
%>
<form id="myform" action="gg_list.asp?id=<%=Request("id")%>&amp;action=ok&amp;ac=ok" method="post" runat="server">
 
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#33CCFF" class="table" id="ggxq">
  <tr>
		<td height="22" colspan="2" align="left">选择申请站点:	  </td>
	</tr>
	
	<%Set rsa= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ad where id="&Request("id")&""
rsa.Open sql,conn,1,2
dim ggidd
ggidd=rsa("id")
%>

<input name="money" type="hidden" value="<%=rsa("money")%>"/>
<input name="ggtitle" type="hidden" value="<%=rsa("title")%>"/>
<input name="ggid" type="hidden" value="<%=ggidd%>"/>
<input name="gglx" type="hidden" value="<%=rsa("gglx")%>"/>
	<tr>
		<td width="20%" height="22" align="left">&nbsp;广告:</td>
		<td width="80%" height="22" align="left">
		<%=rsa("title")%>
		<%rsa.close
		set rsa=nothing%>
		</td></tr>
 	<tr>
		<td width="20%" height="22" align="left">&nbsp;网站:</td>
		<td width="80%" height="22" align="left">
		
      <span id="Label_Sm"> 
	  <select name='url' value='网站选择'></span>
        <%
rs.MoveFirst
While Not rs.EOF
j=j+1
%>
        <option value='<%=rs("id")%>'><%=rs("title")%></option>
        <%
rs.MoveNext
Wend
rs.close
set rs=nothing
%>
      </select></td>
	</tr> 
<tr>
<td height="22" colspan="2" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Button1" value="确认申请" /></td>
	</tr>
</table>
 </form>
 
<%if Request("ac")="cc" then
Response.write "<br/>您已申请的站点：<br/>"
Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggurl where username='"&Request.Cookies("username")&"' and ggid="&Request("id")&" Order By id"
rs.Open sql,conn,1,2
j=0
rs.MoveFirst
While Not rs.EOF
j=j+1
response.write ""&j&"."&rs("urltitle")&"<br/>"
dim ztt
if rs("zt")="1" then
ztt="未审核"
elseif rs("zt")="2" then
ztt="审核通过"
elseif rs("zt")="3" then
ztt="审核不通过"
elseif rs("zt")="4" then
ztt="广告已回收"
end if
Response.write "状态:"&ztt&"<br/>"
rs.MoveNext
Wend
rs.close
set rs=nothing
end if%> 
 
 <br><div class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="gg_list.asp?id=<%=Request("id")%>">返回上级</a></div>
<%
else
Response.write "<p class='hongse'>你还没有添加站点或站点未审核！</p>"
Response.write "<p class='px'><img src='images/fanhui.gif' width='16' height='9'  alt='网站管理'/><a href='url.asp'>网站管理</a></p>"
end if
end if
%>

<%

if Request("ac")="ok" then
if Request("url")="" then
Response.write "请先选择申请站点!"
else
Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT urlid From ggurl where urlid="&Request("url")&" and ggid="&Request("ggid")&""
rs.Open sql,conn,1,2
if not rs.eof then
Response.write "该站点已申请此广告！<br/>"
else
Set rsa= Server.CreateObject("ADODB.Recordset")
sql="SELECT url,id,title From url where id="&Request("url")&""
rsa.Open sql,conn,1,2

Set rsg= Server.CreateObject("ADODB.Recordset")
sql="SELECT ggsh From ad where id="&Request("id")&""
rsg.Open sql,conn,1,2

Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggurl"
rs.Open sql,conn,1,2
rs.addnew
rs("ggid")=Request("ggid")
rs("money")=Request("money")
rs("ggtitle")=Request("ggtitle")
rs("gglx")=Request("gglx")
rs("urltitle")=rsa("title")
rs("url")=Replace(rsa("url"),"http://","")
rs("urlid")=rsa("id")
rs("username")=Request.Cookies("username")
rs("userid")=Request.Cookies("id")
if not rsg.eof then
if rsg("ggsh")=1 then
rs("zt")=2
attea=1
end if
end if
rsa.close
set rsa=nothing
rs.Update
rs.close
set rs=nothing
rsg.close
set rsg=nothing
if attea=1 then
Response.write "<p>广告申请成功，系统已自动审核广告！</p>"
else
Response.write "<p>广告申请成功，请等待管理员审核！</p>"
end if
end if
end if
end if
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="gg.asp">广告列表</a></p>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
