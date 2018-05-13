<!--#INCLUDE file="conn.asp"-->
<title>结算管理</title>
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

<!--结算管理页面-->
<%
  if Request("action")="" then%>
  <div class="p13">&nbsp;结算管理</div>
<% dim page

    		Page = Request("page")      '当前页数
    		If Page = "" Then Page = 1
   		If Not IsNumeric(Page) Then Page = 1
    		Page = CLng(Page)
    		If Page < 1 Then Page = 1  
			
 set rs=Server.CreateObject("ADODB.Recordset")
 if Request("ac")="" then
  sql="select * from jsmoney where zt=0 Order By time desc"
  elseif Request("ac")="yjs" then
   sql="select * from jsmoney where zt=1 Order By time desc"
   end if
  rs.open sql,conn,1,2
  
 %><p><%if Request("ac")="" then%>
<span class="tab"><span>&nbsp;&nbsp;未结算</a>&nbsp;&nbsp;</span><a href="jsgl.asp?ac=yjs">已结算</a></span>
<%elseif Request("ac")="yjs" then%>
<span class="tab">&nbsp;&nbsp;<a href="jsgl.asp">未结算</a><span>&nbsp;&nbsp;已结算</span></span>
<%end if%>
<%
  if rs.eof then
  Response.write "<p>暂无结算申请数据！</p>"
  %> <p><img src="../images/fanhui.gif"/><a href="jsgl.asp">结算管理</a></p>
  <p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
 <%
  Response.end
  end if
   rs.Move((page-1)*ada) 
  dim i
  i=1
%> 
<table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">	
<tr></tr></table>
	<%
	
if not (rs.bof and rs.eof)  then
				Rs.PageSize = 5	'一页N条记录
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
 <table id="ggxq" width="100%" border="1" cellspacing="0" cellpadding="0" align="center" class="table">					
<tr>
<%set rsz=Server.CreateObject("ADODB.Recordset")
   sql="select khh,khdc,khm,yhzh from username where username='"&rs("username")&"'"
  rsz.open sql,conn,1,2%>
用户：<%=rs("username")%><br/>申请金额：<%=rs("money")%><br/>开户银行：<%=rsz("khh")%><br/>
开户地址：<%=rsz("khdc")%><br/>开户姓名：<%=rsz("khm")%><br/>银行帐号：<%=rsz("yhzh")%><br/>
申请时间：<%=rs("time")%><br/>
<%if rs("zt")="1" then%>
汇款时间：<%=rs("timee")%><br/>
支付状态：已结算<br/><p>
<%else%>
支付状态：未结算<br/>
<p><a href="jsgl.asp?action=js&amp;id=<%=rs("id")%>" class="STYLE1">确定结算</a>&nbsp;
<%end if%>
<a href="jsgl.asp?action=sc&amp;id=<%=rs("id")%>" class="STYLE1">删除</a></p>
</tr>
</table>
<% 
	 	Rs.MoveNext
	  	Next

    		'分页
    		If Page < rs.PageCount Then
    		    Response.Write("<a href='?page=" & Page + 1 & "&amp;action="&Request("action")&"&amp;ac="&Request("ac")&"'>下一页</a>")
    		End If
       
    		If Page > 1 And Page < rs.PageCount Then
    		    Response.Write("|")
    		End If
        
    		If Page > 1 Then
    		    Response.Write("<a href='?page=" & Page - 1 & "&amp;action="&Request("action")&"&amp;ac="&Request("ac")&"'>上一页</a><br/>")
    		Elseif rs.PageCount >1 then
    		    Response.Write("<br/>")
    		End If
            
    		Randomize()
    
    		If rs.PageCount > 2 Then
    		    Response.Write("<input name='Page' format='*N' size='5' maxlength='5'/>")
    		    Response.Write("<anchor>跳页")
    		    Response.Write("<go href='?action="&Request("action")&"&amp;ac="&Request("ac")&"' method='post'>")
    		    Response.Write("<postfield name='Page' value='$(Page:n)' />")
    		    Response.Write("</go></anchor><br/>")
    		End If
     		    Response.Write("[第"&Page&"/总"&rs.PageCount&"页/"&rs.RecordCount&"条]<br/>")
		end if
		Rs.close
	set rs=nothing
end if
%>

 <!-- 结算申请删除页面-->
 <%if Request("action")="sc" then %>
  <div class="p13">&nbsp;结算管理</div>
 <% if Request("ac")="" then
 response.write "<p>您确定删除用户申请记录吗?<br/><a href='jsgl.asp?id="&Request("id")&"&amp;action=sc&amp;ac=ok'>确定</a>&nbsp;&nbsp;<a href='jsgl.asp'>取消</a></p>"
 else
 set rs=Server.CreateObject("ADODB.Recordset")
  sql="delete * from jsmoney where id="&Request("id")&""
  conn.execute sql
 Response.write "<p>结算申请删除成功！</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='jsgl.asp'>结算管理</a></p>"
 end if
 end if
%> 

<!--结算页面-->
<%if Request("action")="js" then%>
 <div class="p13">&nbsp;结算管理</div>
<% 
if Request("ac")="" then
Response.write "<p><a href='jsgl.asp?action=js&amp;ac=ok&amp;id="&Request("id")&"' class='STYLE1''>确定支付</a></p>"
 else
 
 set rs=Server.CreateObject("ADODB.Recordset")
    sql="select zt,timee,money,username from jsmoney where id="&Request("id")&""
  rs.open sql,conn,1,2
set rsa=Server.CreateObject("ADODB.Recordset")
    sql="select * from cwrz"
  rsa.open sql,conn,1,2  
  rsa.addnew
  rsa("username")=rs("username")
  rsa("money")=rs("money")
  moneyt=rs("money")
  rsa("sm")="给用户"&rs("username")&"结算"&rs("money")&"元"
  rsa.update
  rsa.close
  set rsa=nothing
rs("zt")=1
rs("timee")=now()
rs.Update

''-------给下线冲值提成金额---------
 set rsz=Server.CreateObject("ADODB.Recordset")
    sql="select sxuser,username from username where username='"&rs("username")&"'"
  rsz.open sql,conn,1,2
  if not rsz.eof then
  if rsz("sxuser")<>"" then
  set rsb=Server.CreateObject("ADODB.Recordset")
    sql="select sxuser,username,money from username where username='"&rsz("sxuser")&"'"
  rsb.open sql,conn,1,2
  rsb("money")=rsb("money")+moneyt*0.05
  rsb.Update
  rsb.close
  set rsb=nothing
  set rsa=Server.CreateObject("ADODB.Recordset")
    sql="select * from cwrz"
  rsa.open sql,conn,1,2  
  rsa.addnew
  rsa("username")=rsz("sxuser")
  moneyz=moneyt*0.05
  rsa("money")=moneyz
  rsa("sm")="给用户"&rs("username")&"的下线会员"&rsz("sxuser")&"冲值结算提成"&moneyz&"元"
  rsa.update
  rsa.close
  set rsa=nothing
  end if
  end if
  rsz.close
  set rsz=nothing
rs.close
set rs=nothing

  
Response.write "<p>结算成功！<br/>结算下线提成功</p>"
 Response.write "<p><img src='../images/fanhui.gif'/><a href='jsgl.asp'>结算管理</a></p>"
end if
%>
<%end if%>

<!--生成汇款单-->
<%if Request("action")="hkd" then%>
  <div class="p13">&nbsp;生成汇款单</div>
<% set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from jsmoney where zt=0 order by id desc"
  rs.open sql,conn
  dim content
  
  do while not rs.eof
set rsz=Server.CreateObject("ADODB.Recordset")
   sql="select khh,khdc,khm,yhzh from username where username='"&rs("username")&"'"
  rsz.open sql,conn
                content=content & "商户:" & rs("username")  & vbNewLine  
                content=content & "开户银行:" & rsz("khh")  & vbNewLine
				content=content & "开户名:" & rsz("khm")  & vbNewLine
				content=content & "银行帐号:" & rsz("yhzh")  & vbNewLine
				content=content & "开户地址:" & rsz("khdc")  & vbNewLine
				content=content & "支付金额:"&rs("money")&"元" & vbNewLine
				content=content & "------------------" & vbNewLine
  
rs.movenext
   loop
   
   dim FSO,TS

		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		IF FSO.FileExists(Server.MapPath("./hkd.txt")) = True then
   		Fso.deleteFile Server.MapPath("./hkd.txt"),true
		end if

    		Set FSO = Server.CreateObject("Scripting.FileSystemObject")   
    		Set TS = FSO.OpenTextFile(Server.MapPath("./hkd.txt"),8,true) 
  		TS.write content   
    		Set TS = Nothing   
    		Set FSO = Nothing   
   
rs.close
set rs=nothing
%>生成汇款单成功!<br/><p><a href='hkd.txt'>下载汇款单</a></p>----------<br/>
<%end if%>
 
<p><img src="../images/fanhui.gif"/><a href="jsgl.asp?action=hkd" class="STYLE1">生成汇款单</a></p>
<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>

