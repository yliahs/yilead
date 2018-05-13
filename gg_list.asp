<!--#INCLUDE file="conn.asp"-->
<title>广告详情</title>	

</head>
<body>

    <div id="all">
    
<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>
<%if Request("action")="" then%>
      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />广告代码</div>
  <div></div>
  
  <%
   set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From ad where id="&Request("id")&""
  rs.open sql,conn
  dim gglx,title,money,jmoney,ggy,ggsm,gid,logo
  if rs("gglx")="1" then
  gglx="点击类广告："
  else
  gglx="效果类广告："
  end if
  title=rs("title")
  money=rs("money")
  jmoney=rs("jmoney")
  ggy=rs("ggy")
  ggsm=rs("ggsm")
  gid=rs("id")
  logo=rs("logo")
  rs.close
  set rs=nothing
  Set rs= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggurl where ggid="&Request("id")&" and username='"&Request.Cookies("username")&"'"
rs.Open sql,conn,1,2
  %>
  <p class="hongse"><%=gglx%></p>
  
        
    <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">
	<tr>
		<td width="22%" height="22" align="center">&nbsp;LOGO</td>
		<%if logo<>"" then%>
		<td height="50">
      <span id="Label_GGy"><img src="<%=logo%>"/></span>	
	  <%else%>
		<td height="22">
      <span id="Label_GGy">暂无广告图片</span>
	  <%end if%>
      </td>
	</tr>
	<tr>
		<td width="22%" height="22" align="center">
          <span id="Label1" style="color:#000F00;">名称</span> </td>
		<td width="78%" height="22">&nbsp;<span id="Label_title" style="color:#000F00;"><%=title%></span></td>
	</tr>
	<tr>
		<td width="22%" height="22" align="center">
          <span id="Label2" style="color:#000F00;">价格</span> </td>
		<td width="78%" height="22">&nbsp;<span id="Label_Price" style="color:#000F00;"><%=money%>元/<%=jmoney%></span></td>
	</tr>
	<tr>
		<td height="22" align="center">&nbsp;广告语</td>
		<td height="22">
      <span id="Label_GGy"><p><%=ggy%></p></span>
      </td>
	</tr>
    <%if not (rs.eof and rs.bof) then%>
    <%if rs("zt")="2" then%>
	<tr id="Tr_ggUrl">
		<td height="35" align="center">&nbsp;代　码</td>
		<td height="35">
      &nbsp;<span id="Label_dm1" style="font-size:13px; color:Blue ;"><%=session("url")%>/c?c=<%=Request.Cookies("id")%>@<%=gid%></span>
      </td>
	</tr>
    <%else%>
    <%dim zt
	if rs("zt")="1" then
	zt="广告未审核"
	elseif rs("zt")="3" then
	zt="广告审核未通过"
	elseif rs("zt")="4" then
	zt="广告已回收"
	end if
	%>
    <tr id="Tr_ggUr3">
		<td height="22" align="center">&nbsp;状态</td>
		<td height="22">
      &nbsp;<span id="Label_dm1" style="font-size:13px; color:#FF0000"><%=zt%></span>
      </td>
	</tr>
    <%end if%>
    <%end if%>
	<tr>
		<td height="22" align="center">&nbsp;说明</td>
		<td height="22">
      <span id="Label_Sm"><p><%=ggsm%></p></span>
      </td>
	</tr>
	<tr>
		<td height="22" colspan="2" align="center">
        <%if not(rs.eof and rs.bof) then%>
        <%if rs("zt")<>"4" then%>
          <a id="HyperLink1" href="gg_list.asp?id=<%=Request("id")%>&amp;action=ok&amp;ac=cc">管理网站</a>
          <%else%>
          <a id="HyperLink1" href="gg_list.asp?id=<%=Request("id")%>&amp;action=ok">申请广告</a>
          <%end if%>
          <%else%>
           <a id="HyperLink1" href="gg_list.asp?id=<%=Request("id")%>&amp;action=ok">申请广告</a>
          <%end if%>
      </td>
	</tr>
</table>
<%
rs.close
set rs=nothing
%>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='广告列表'/><a href="gg.asp">广告列表</a></p>
<p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
<%end if%>
<%if Request("action")="ok" then%>
<!--#INCLUDE file="gg_url.asp"-->
<%end if%>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
