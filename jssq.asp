<!--#INCLUDE file="conn.asp"-->
<title>结算申请</title>	
</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>
<%if Request("action")="" then%>
      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />结算申请</div>
 
 <%
 
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From username where username='"&Request.Cookies("username")&"'"
  rs.open sql,conn
 %>
<p>&nbsp;您当前金额:<%=FormatNumber(rs("money"),3,-1,-1,0)%>元</p> 
<%
  set rsz=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where id=1"
  rsz.open sql,conn,1,2
  if rsz("sfz")=2 then
  set rssz=Server.CreateObject("ADODB.Recordset")
  sql="select * from sfz where username='"&Request.Cookies("username")&"'"
  rssz.open sql,conn,1,2
  if rssz.eof then
  Response.write "<p>您没有通过身份证件认证，不能申请结算！</p>"
  else
  if rssz("zt")=2 then
  call js
  else
  Response.write "<p>您没有通过身份证件认证，不能申请结算！</p>"
  end if
  end if
  else
  call js
  end if 
  rsz.close
  set rsz=nothing
  %>
 
  <%Function js  
if rs("money")<10 then%>
<br/>您的余额不足10元,请到10元后再申请。<br/>
<%else%>
 <form name="form1" method="post" action="jssq.asp?action=ok">
<br/>申请金额：
  <input name="money" type="text" id="TextBox1" style="width:120px;" size="4" maxlength="5" />
  元<br/>
<input type="submit" name="Button1" value="确认申请" />
  </form>
<%end if%>
<%rs.close
set rs=nothing
End Function%>
<%end if%>


<%if Request("action")="ok" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />结算申请</div>
<%dim money
money=Request("money")
if cint(money)<10 then
Response.write "<br/>申请金额不可小于10元！<br/><br/><a href='jssq.asp'>返回重新输入</a>"
else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT money From username where username='"&Request.Cookies("username")&"'"
  rs.open sql,conn
if rs("money")<cint(money) then
Response.write "<br/>你的金额不足"&money&"元,不能申请！<br/>"
else
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT money From username where username='"&Request.Cookies("username")&"'"
  rs.open sql,conn,1,2
  rs("money")=cint(rs("money")-cint(money))
  rs.Update
  rs.close
  set rs=nothing
  
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From jsmoney"
  rs.open sql,conn,1,2
  rs.addnew
  rs("money")=money
  rs("username")=Request.Cookies("username")  
  rs.Update
  rs.close
  set rs=nothing
Response.write "<br/>&nbsp;申请成功!<br/>"
Response.write "<br/><img src='images/fanhui.gif' width='16' height='9'/><a href='jssq.asp?action=jl'>结算记录</a>"
end if
end if
%>
<%end if%>


<%if Request("action")="jl" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />结算记录</div>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From jsmoney where username='"&Request.Cookies("username")&"' Order By time desc"
  rs.open sql,conn,1,2

%>
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
<tr>       
		<td width="25%" height="22" align="center">&nbsp;申请金额</td>
		<td width="55%" height="22" align="center">&nbsp;申请时间</td>
        <td width="20%" height="22" align="center">&nbsp;状&nbsp;态</td>
	</tr>
  <% if not(rs.eof and rs.bof) then
  rs.MoveFirst
        While Not rs.EOF
  dim zt
  if rs("zt")="0" then
  zt="未结算"
  else
  zt="已结算"
  end if
   %>
    <tr>
		<td width="25%" height="22" align="center">&nbsp;<%=rs("money")%>元</td>
		<td width="55%" height="22" align="center">&nbsp;<%=rs("time")%></td>
               <td width="20%" height="22" align="center">&nbsp;<%=zt%></td>
	</tr>
    <%
	rs.MoveNext
    Wend
	%>
  <%else%>
   <tr>
		<td height="22" colspan="3" align="center">您没有结算记录！</td>
        </tr>
   <%end if%> 
</table>
<%end if%>

 <p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
