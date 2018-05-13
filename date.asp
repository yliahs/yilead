<%
dbb="/#date/yilead.mdb" 
Provider="Provider=Microsoft.Jet.OLEDB.4.0;"
DBPath="Data Source="&Server.MapPath(""&dbb&"")
On Error Resume Next
Set conn=Server.CreateObject("ADODB.connection")
conn.Open Provider&DBPath
If Err Then
		err.Clear
		Set conn = Nothing
		Response.Write "数据库连接出错，请检查连接字符串。</wml>"
		Response.End
	End If
If Err Then
err.Clear
end if%>