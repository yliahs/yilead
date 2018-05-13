<!--#INCLUDE file="conn.asp"-->
<!--#INCLUDE file="../hs.asp"-->
<title>数据管理</title>
<style type="text/css">
<!--
.STYLE1 {color: #0000FF}
-->
</style>
</head>
<body>
<div id="all">   
 
<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
<%
'select case action
'case ""
'    call inde
'case "yjrl"
'	call yjrl
'case "2"
'	call HeadShow1
'case "3"
'	call Recommend0
'case "Recommend1"
'end select
if Request("action")="" then
call inde
elseif REquest("action")="yjrl" then
call yjrl
elseif Request("action")="gg" then
call gg
elseif Request("action")="sjrl" then
call sjrl
end if 
%>


<% sub inde %>
  <div class="p13">&nbsp;数据录入</div>
  <p><a href="adsj.asp?action=yjrl" class="STYLE1">一键自动批量录入</a><br/>
<a href="adsj.asp?action=gg" class="STYLE1">按广告手动录入</a><br/>------------</p>
 <% if Request("ac")="" then%> 
<span class="tab"><span>点击计费类</a></span>&nbsp;<a href="adsj.asp?ac=yjs">效果计费类</a></span>
<%else%>
<span class="tab">&nbsp;<a href="adsj.asp">点击计费类</a>&nbsp;<span>效果计费类</span></span>
<%end if%>
<br/>-----------<br/>
<form id="myform" action="adsj.asp?acc=ad" method="post" runat="server">
<input name="time" type="text" value="<%=date-1%>" size="10"/>
<input name="ac" type="hidden" value="<%=Request("ac")%>"/>
<input type="submit" name="Button1" value="查询" />
</form><br/>
 <%

 call az

end sub%>


 <%sub az
  timee=Request("time")
 if timee="" then
 timee=date-1
 end if
 Response.write "<p>日期："&timee&"</p>"
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("ac")="" then
  sql="select * from ggsj where time=#"&timee&"# and gglx='1' and zt=0 Order By zrdjip desc"
   elseif Request("ac")="yjs" then
   sql="select * from ggsj where time=#"&timee&"# and gglx='2' and zt=0 Order By zrdjip desc"
  elseif Request("ac")="aa" then
  sql="select * from ggsj where time=#"&timee&"# and ggid='"&Request("id")&"' and zt=0 Order By zrdjip desc"
   end if
  rs.open sql,conn,1,2
 if rs.eof then
 Response.write "<p>暂无数据!</a><br/>&nbsp;<a href='index.asp'>返回管理首页</a><br/>"
 Response.end
 end if
    rs.Move((page-1)*ada) 
  dim i
  i=1 
  if not (rs.bof and rs.eof)  then
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
	
Response.write "<br/>&nbsp;广告名称："&rs("ggtitle")&"<br/>&nbsp;站长帐号："&rs("username")&"<br/>&nbsp;点击 IP："&rs("zrdjip")&"<br/>&nbsp;点出PV："&rs("zrdjpv")&"<br/>"
if Request("ac")="aa" then
Response.write "<a href='adsj.asp?action=sjrl&amp;time="&timee&"&amp;username="&rs("username")&"&amp;id="&Request("id")&"' class='STYLE1'>&nbsp;数据录入</a><br/>---------<br/>"
else
Response.write "<a href='adsj.asp?action=sjrl&amp;time="&timee&"&amp;username="&rs("username")&"&amp;id="&rs("ggid")&"' class='STYLE1'>&nbsp;数据录入</a><br/>---------<br/>"
end if
		Rs.MoveNext
	  	Next
    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='url.asp?page=" & Page + 1 & "&amp;action="&Request("action")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='url.asp?page=" & Page - 1 & "&amp;action="&Request("action")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='url.asp?action="&Request("action")&"' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		rs.close
		set rs=nothing

end sub%>


<%Sub yjrl %>
<div class="p13">&nbsp;一键批量录入</div>
<%if Request("ac")="" or Request("ac")="yjs" then%>
<% if Request("ac")="" then%> 
<span class="tab"><span>点击计费类</a></span>&nbsp;<a href="adsj.asp?action=yjrl&amp;ac=yjs">效果计费类</a></span>
<%else%>
<span class="tab">&nbsp;<a href="adsj.asp?action=yjrl">点击计费类</a>&nbsp;<span>效果计费类</span></span>
<%end if%><br/><br/>
<p>一键站长数据自动录入</p>----------<br/>

<%set rs=Server.CreateObject("ADODB.Recordset")
 if Request("ac")="" then
  sql="select * from ad where gglx=1 Order By id desc"
  elseif Request("ac")="yjs" then
  sql="select * from ad where gglx=2 Order By id desc"
   end if
  rs.open sql,conn,1,2
  dim ii
  ii=0
  %> <form id="myform" action="adsj.asp?action=yjrl&amp;ac=ad" method="post" runat="server">
 &nbsp;日期：<input name="time" type="text" value="<%=date-1%>" size="11"/>
 <%
do while not rs.eof
ii=ii+1%>
  <% Response.write "<p>广告："&rs("title")&"<br/>"%>
  扣量百分比：
<input name="kl<%=ii%>" type="text" value="30" size="4"/>
%<br/>
<input name="ggid<%=ii%>" type="hidden" value="<%=rs("id")%>"/>
<input name="money<%=ii%>" type="hidden" value="<%=rs("money")%>"/>
<%rs.movenext 
loop %>
<input name="ii" type="hidden" value="<%=ii%>">
<input type="submit" name="Button1" value="确定一键录入" />
</form><br/></p>
<%rs.close
set rs=nothing%>
<%end if%>

<%if Request("ac")="ad" then
for i=1 to Request("ii")
  dim aaa,zt
  aaa=1
  zt=0
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggsj where ggid='"&Request("ggid"&i)&"' and zrdjip>"&aaa&" and zt="&zt&" and time=#"&Request("time")&"#"
  rs.open sql,conn,1,2
  
  if not rs.eof then
   do while not rs.eof
  set rss=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggsj where username='"&rs("username")&"' and zt=0 and time=#"&Request("time")&"#"
  rss.open sql,conn,1,2
  zrdjip=cint(rs("zrdjip")-rs("zrdjip")*(Request("kl"&i)/100))
  rss("zrsj")=zrdjip
  rss("zt")="1"
  rss.update
  rss.close
  set rss=nothing

set rsv=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggfw where ggid='"&Request("ggid"&i)&"' and time=#"&Request("time")&"#"
  rsv.open sql,conn,1,2
  rsv("tzsj")=cint(rsv("tzsj")+zrdjip)
  rsv.update
  rsv.close
  set rsv=nothing
   
  set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select * from username where username='"&rs("username")&"'"
  rsa.open sql,conn,1,2
  monaa=""&zrdjip*Request("money"&i)&""
  if monaa<1 then
  monaa="0"&monaa&""
  end if
  rsa("money")=rsa("money")+monaa
  rsa.update
  rsa.close
  set rsa=nothing
  
  set rsq=Server.CreateObject("ADODB.Recordset")
  sql="select title from ad where id="&Request("ggid"&i)&""
  rsq.open sql,conn,1,2
  dim mone
  mone=zrdjip*Request("money"&i)
  if mone<1 then
  mone="0"&mone&""
  end if
    jysmm="给用户"&rs("username")&"广告"&rsq("title")&"添加"&zrdjip&"条数据"
  call money(rs("username"),mone,1,jysmm)
  rsq.close
  set rsq=nothing
  RecordCount=rs.RecordCount
  rs.movenext 
  loop 
  end if
  rs.close
  set rs=nothing
next
if RecordCount="" then
RecordCount=0
end if 
Response.write "<p>全部数据录入成功!<br/>共给"&RecordCount&"位站长录入数据<p>"  
end if%>
<p><img src="../images/fanhui.gif"/><a href="adsj.asp">数据录入</a></p>
<%End Sub%>


<%sub gg %>
 <div class="p13">&nbsp;数据录入</div>
 <%if Request("ac")="" then%>
<%set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad Order By id desc"
  rs.open sql,conn,1,2
  q=0
  if not rs.bof then
do while not rs.eof
q=q+1
Response.write "<p>"&q&".<a href='adsj.asp?action=gg&amp;ac=aa&amp;id="&rs("id")&"'>"&rs("title")&"</a></p>"
rs.movenext
loop
rs.close
set rs=nothing
else
Response.write "<p>暂无广告！</p>"
end if
end if
if Request("ac")="aa" then%>
 <form id="myform" action="adsj.asp?action=gg&amp;ac=aa&amp;id=<%=Request("id")%>" method="post" runat="server">
 &nbsp;日期：<input name="time" type="text" value="<%=date-1%>" size="10"/>
<input type="submit" name="Button1" value="查询" />
</form><br/></p>

<%call az
end if
end sub%>

<%sub sjrl%>
 <div class="p13">&nbsp;数据录入</div>
<% if Request("ac")="" then
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggsj where time=#"&Request("time")&"# and ggid='"&Request("id")&"' and username='"&Request("username")&"'"
  rs.open sql,conn,1,2
Response.write "<p>用户"&rs("username")&"<br/>点击IP："&rs("zrdjip")&"<br/>点击PV:"&rs("zrdjpv")&"<br/>独立下载："&rs("xzcs")&"<br/>成功安装："&rs("anzcs")&"<br/>有效注册："&rs("yxzc")&"个</p>"%>
<form id="myform" action="adsj.asp?action=sjrl&amp;ac=aa" method="post" runat="server">
 &nbsp;数据：<input name="sj" type="text" value="" size="8"/><br/>
 <input name="time" type="hidden" value="<%=Request("time")%>">
  <input name="id" type="hidden" value="<%=Request("id")%>">
    <input name="username" type="hidden" value="<%=Request("username")%>">
&nbsp;<input type="submit" name="Button1" value="确定录入" />
</form><br/>
<p><img src="../images/fanhui.gif"/><a href="<%=request.ServerVariables("HTTP_REFERER")%>">返回上级</a></p>
<%else%>
<%set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggsj where time=#"&Request("time")&"# and ggid='"&Request("id")&"' and username='"&Request("username")&"'"
  rs.open sql,conn,1,2
  rs("zrsj")=Request("sj")
  rs("zt")=1
  rs.update
  rs.close
  set rs=nothing
  
   
  set rsa=Server.CreateObject("ADODB.Recordset")
  sql="select money,title from ad where id="&Request("id")&""
  rsa.open sql,conn,1,2
  dim moneyy 
  moneyy=Request("sj")*rsa("money")
  title=rsa("title")
   rsa.close
  set rsa=nothing
  jysm="给用户"&Request("username")&"广告"&title&"添加"&Request("sj")&"条数据"
   if moneyy<1 then
  moneyy="0"&moneyy&""
  end if
  
 call money(Request("username"),moneyy,1,jysm)
  
  Response.write "<p>数据录入成功</p>"
  Response.write "<p>总计支付："&moneyy&"元</p>"
 %> <p><img src="../images/fanhui.gif"/><a href="adsj.asp?action=gg&amp;time=<%=Request("time")%>&amp;ac=aa">返回上级</a></p>
<%end if%>

<p><img src="../images/fanhui.gif"/><a href="adsj.asp">数据录入</a></p>
<%end sub%>

<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>