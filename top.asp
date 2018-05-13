<%set objgbrs=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&Request.Cookies("username")&"'"
  objgbrs.open sql,conn
  if objgbrs("password")<>Request.Cookies("password") then
  Response.Redirect "login.asp"
  Response.end
  else %>
<p><div class="p31">&nbsp;合作ID:<%=Request.Cookies("id")%>|<a href="pc.asp">个人中心</a>|<a href="sms.asp">信箱</a>|<a id="Top1_HyperLink1" href="login.asp?action=zx">注销</a></div></p>

<%  end if
  objgbrs.close
  set objgbrs=nothing
%>