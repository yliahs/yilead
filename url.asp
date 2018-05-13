<!--#INCLUDE file="conn.asp"-->
<title>网站管理</title>	
<style type="text/css">
<!--
.STYLE1 {color: #0000FF}
.STYLE2 {color: #FF0000}
-->
</style>
</head>
<body>
<div id="all">
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

  <%if Request("action")="" then%>
  
      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />网站管理</div>
	  <p align="center"><a href="url.asp?action=ok"><span id="Label_dm1" style="font-size:13px; color:Blue ;">增加站点</span></a></p>
  <%
  Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From url where username='"&Request.Cookies("username")&"' Order By id"
rs.Open sql,conn,1,2
%>

 <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#33CCFF" class="table" id="ggxq">
  
	<tr>
		<td width="45%" height="22" align="left">&nbsp;网站名称</td>
		<td width="55%" height="22" align="left">&nbsp;网站地址		</td>
	</tr>
		<%
		if not(rs.eof and rs.bof) then
		rs.MoveFirst
        While Not rs.EOF%>
 	<tr>

	   <td width="45%" height="22" align="left">&nbsp;<%=rs("title")%>&nbsp;<a href="url.asp?id=<%=rs("id")%>&amp;action=sc" class="STYLE1">删除</a></td>

	   <%if rs("zt")="1" then%>
<td width="55%" height="22" align="left">&nbsp;<span class="STYLE2">站点未审核</span></td>
 	</tr>		
		  <%elseif rs("zt")="3" then%>
	<td width="45%" height="22" align="left">&nbsp;<span class="STYLE2">审核不通过</span></td>
	</tr>		  
		  <%elseif rs("zt")="2" then%>
		  <%if Instr(rs("url"),"http://") then%> 
<td width="45%" height="22" align="left">&nbsp;<a href="<%=rs("url")%>"><%=rs("url")%></a></td></tr>
<%else%>
<td width="45%" height="22" align="left">&nbsp;<a href="http://<%=rs("url")%>">http://<%=rs("url")%></a></td></tr>
<%end if%>
		  	  <%end if%>
        <%
rs.MoveNext
Wend
else
%><tr><td height="22" colspan="2" align="center">您还没有添加站点！</td></tr><%
end if
rs.close
set rs=nothing
%>
</table>
<%end if%>

<%if Request("action")="ok" then%>
<%if Request("ac")="" then%>
 <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />添加站点</div>
 <form name="form1" method="post" action="url.asp?action=ok&amp;ac=ok">

   <p>网站名称:<br/>
      <input name="urltitle" type="text" maxlength="20" id="TextBox1" style="width:120px;" />
      <br/>
     网站地址:<br/>
      <input name="url" type="text" id="TextBox1" style="width:120px;" value="http://" maxlength="40" />
	<br/> 网站类型:<br/>
	  <select name="urllx" id="DropDownList1">
		<option selected="selected" value="门户">门户</option>
		<option value="社区">社区</option>
		<option value="网址">网址</option>
		<option value="下载">下载</option>
		<option value="书刊">书刊</option>
		<option value="两性">两性</option>
		<option value="彩票">彩票</option>
		<option value="行业">行业</option>
		<option value="企业">企业</option>
		<option value="个人博客">个人博客</option>
		<option value="资讯">资讯</option>
		<option value="动漫">动漫</option>
		<option value="汽车">汽车</option>
		<option value="手机">手机</option>
		<option value="图片">图片</option>
		<option value="其它">其它</option>

	</select>
    </p>
	<input type="submit" name="Button1" value="确认注册" />
  </form>
 <%end if%>
 <%if Request("ac")="ok" then%>
 <%dim urltitle,url,urllx
 urllx=Request("urllx")
 urltitle=Request("urltitle")
 url=Request("url")
 url=Replace(url,"http://","")
url=split(trim(url),"/")
 if urltitle="" then daving=daving&"<br/>网站名称不可为空！<br/>"
 if Request("url")="" then daving=daving&"<br/>网站地址不可为空！<br/>"
 if Request("url")="http://" then daving=daving&"<br/>网站地址不可为空！<br/>"
 if urltitle="" or Request("url")="" then
 Response.Write daving
 else
 Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From url where url='"&trim(url(0))&"'"
rs.Open sql,conn,1,2
 if not rs.eof then
 Response.Write "<br/>该站点已存在，请重新添加!<br/><br/><a href='url.asp?action=ok'>返回重写</a><br/>"
 else 
 Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From url"
rs.Open sql,conn,1,2
 rs.addnew
 rs("title")=urltitle
 rs("url")=trim(url(0))
 rs("username")=Request.Cookies("username")
 rs("urllx")=urllx
 rs("zt")=session("ik")
 rs.Update
 rs.close
 set rs=nothing
 Response.write "<br/>添加站点成功！<br/><br/><img src='images/fanhui.gif' width='16' height='9'/><a href='url.asp'>网站管理</a><br/>"
 end if
 
 end if
 %>

<%end if%>
<%end if%>


<%if Request("action")="sc" then%>
<div class="p12"> <img src="images/lalala.gif" width="14" height="14" />删除站点</div>
<%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="delete From url where username='"&Request.Cookies("username")&"' and id="&Request("id")
conn.execute sql
Response.write "<br/>站点删除成功!<br/><br/><img src='images/fanhui.gif' width='16' height='9'/><a href='url.asp'>网站管理</a><br/>"
end if%>

<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>