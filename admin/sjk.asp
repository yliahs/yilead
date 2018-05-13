<!--#INCLUDE file="conn.asp"-->
<title>数据库管理</title>
</head>
<body>
<div id="all">   

<%if Request.Cookies("admin")="" then
Response.Redirect "login.asp"
Response.end
else%>
<br/>管理员:<%=Request.Cookies("admin")%>&nbsp;|&nbsp;<a id="Top1_HyperLink1" href="login.asp?action=zx">注销登陆</a><br/>
<%end if%>
请选择以下操作：<br/><a href="sjk.asp?action=bf">备份数据库</a>/<a href="sjk.asp?action=fy">还原数据库</a><br/>
<%if Request("action")="fy" then
call fy
elseif Request("action")="bf" then
call bf
end if
%>

<%Function bf%>
<div class="p13">&nbsp;备份数据库</div> 
<%if Request("ac")="" then%>
<p>请输入数据库的相绝对路径,不输入目录就会备份到主目录:<br/>
	注意:如果输入路径中包含文件夹,请确认存在该文件夹,否则备份出错!</p>
	<form action="sjk.asp?action=bf&amp;ac=ok" method="post">
	<p>备份数据库路径：<br/>
	<input emptyok="true" name="sjk" type="text"  maxlength="50" value="<%=date()%>.bak" /><br/>
	<input name="" type="submit" value="开始备份">
	</p>
	</form>
<%
else
if fsolimit=true then
Response.write "<p>注意：此功能需要FSO的支持，您现在使用的服务器<b>不支持</b>该组件,所以该功能无法实现！</p>"
else
'Response.write "<p>注意：此功能需要FSO的支持，您现在使用的服务器<b>支持</b>该组件。</p>"
if Request("sjk")="" then 
Response.write "<p>数据库的相绝对路径不能为空</p>"
else
sjk=server.mappath("../" & trim(Request("sjk")))
		Dim fso, Engine, strDBPath,JET_3X 
		strDBPath = left(sjk,instrrev(sjk,"\")) 
		Set fso = CreateObject("Scripting.FileSystemObject") 

		If fso.FileExists(server.mappath(""&dbb&"")) Then 
			fso.copyfile server.mappath(""&dbb&""),sjk

			%><p><font color="#FF0000">你的数据库已经备份成功!</font></p><%

		Else 
			%><p><font color="#FF0000">数据库名称或路径不正确. 请重试!或用备份数据库备份!</font></p><%
		End If 
end if
end if
end if%>
<%End Function%>


<%Function fy%>
<div class="p13">&nbsp;还原数据库</div> 
<%if Request("ac")="" then%>
<p>请输入数据库的相绝对路径,不输入目录就会从主目录还原:<br/>
	注意:如果输入路径中包含文件夹,请确认存在该文件夹,否则还原出错!</p>
<form action="sjk.asp?action=fy&amp;ac=ok" method="post">
	<p>还原数据库路径：<br/>
	<input emptyok="true" name="sjk" type="text"  maxlength="50" value="<%=date()%>.bak" /><br/>
	<input name="" type="submit" value="开始还原">
	</p>
	</form>
<%else
if fsolimit=true then
Response.write "<p>注意：此功能需要FSO的支持，您现在使用的服务器<b>不支持</b>该组件,所以该功能无法实现！</p>"
else
if Request("sjk")="" then 
Response.write "<p>数据库的相绝对路径不能为空</p>"
else
sjk=server.mappath("../" & trim(Request("sjk")))
		'Dim fso, Engine, strDBPath,JET_3X 
		strDBPath = left(sjk,instrrev(sjk,"\")) 
		Set fso = CreateObject("Scripting.FileSystemObject") 

		If fso.FileExists(sjk) Then 
			fso.copyfile sjk,server.mappath(""&dbb&"")

			%><p><font color="#FF0000">你的数据库已经还原成功!</font></p><%

		Else 
			%><p><font color="#FF0000">备份目录下并无您的备份文件！</font></p><%
		End If 

end if
end if
end if%>
<%End Function%>


<%
function fsolimit()
	on error resume next
	fsolimit=false
	dim fsolimitstr,testfso
	Set testfso = CreateObject("Scripting.FileSystemObject") 
	if not isobject(testfso) then	'不支持FSO
		fsolimit=true
	else
		fsolimit=false
	end if	
end function
%>




<p><img src="../images/fanhui.gif"/><a href="index.asp">管理首页</a></p>
</div>
</body>
</html>
