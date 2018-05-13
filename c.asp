<% Response.ContentType="text/vnd.wap.wml; charset=utf-8" %>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="max-age=0"/>
<meta http-equiv="Cache-Control" content="no-cache"/>
<meta http-equiv="Expires" content="Mon, 1 Jan 1990 00:00:00 GMT"/>
</head>
<%
Provider="Provider=Microsoft.Jet.OLEDB.4.0;"
DBPath="Data Source="&Server.MapPath("ad.mdb")
Set conn=Server.CreateObject("ADODB.connection")
conn.Open Provider&DBPath
%>
    <%  '判断是广告地址1进来的还是广告地址2进来的
  If Instr(Request("c"),"@") Then
  size=trim(Request("c"))
  myarr=split(size,"@")
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&trim(myarr(0))&"' and ggid="&trim(myarr(1))&""
  rs.open sql,conn
  else
set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where id="&cint(Request("c"))&""
  rs.open sql,conn
  end if
  '判断用户申请的广告状态 
  if not rs.eof then
dim zz
zz=rs("zt")
  if zz="1" then zt=zt&"您申请的广告还没有审核！"
  if zz="3" then zt=zt&"您申请的广告没有审核通过！"
  if zz="4" then zt=zt&"您的广告代码已被系统收回！"
 if zz="1" or zz="3" or zz="4" then
  Response.write "<card id='card1' title='访问错误'>"
  Response.write zt
  rs.close
  set rs=nothing
  Response.write "</card></wml>"
  Response.end
  end if
  else
  Response.write "<card id='card1' title='访问错误'>"
  Response.write "此广告已经下架或广告代码不存在！请您及时更换广告代码，以免造成流量损失！"
  rs.close
  set rs=nothing
  Response.write "</card></wml>"
  Response.end
  end if
  %>
  
  <%'判断来源地址
  set rsaa=Server.CreateObject("ADODB.Recordset")
  sql="select fwxz from admin where id=1"
  rsaa.open sql,conn
  if rsaa("fwxz")=1 then
  dim urls
    urls=request.ServerVariables("HTTP_REFERER")
	if urls="" then
Response.write "<card id='card1' title='访问错误'>"
 Response.write "来源址址：空<br/>来源地址为空不能访问广告,请将广告代码挂到您已申请的站点里面，谢谢合作！<br/>" 
 Response.write "</card></wml>"
 Response.end
end if 
	urls=Replace(urls,"http://","")
	urls=split(trim(urls),"/")
	  If Instr(Request("c"),"@") Then
  size=trim(Request("c"))
  myarr=split(size,"@")
  set rst=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&trim(myarr(0))&"' and ggid="&trim(myarr(1))&" and url='"&trim(urls(0))&"' and zt=2"
  rst.open sql,conn
  else
set rst=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&rs("userid")&"' and ggid="&rs("ggid")&" and url='"&trim(urls(0))&"' and zt=2"
  rst.open sql,conn
  end if
if rst.eof then
  Response.write "<card id='card1' title='访问错误'>"
 Response.write "来源地址："&urls(0)&"<br/>此来源站点没有申请此广告或没有审核通过！<br/>" 
 Response.write "</card></wml>"
 Response.end
 end if
rst.close
set rst=nothing	 
end if
rsaa.close
set rsaa=nothing
  %>
  
    <%
set rss=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where id="&rs("ggid")&""
  rss.open sql,conn
  %>
 
<%
'======屏蔽UA客户端=========
browsers=Lcase(Left(Request.ServerVariables("HTTP_USER_AGENT"),4))
if browsers="oper" or browsers="winw" or browsers="wap/" or browsers="wapi" or browsers="mc21" or browsers="up.b" or browsers="upg1" or browsers="upsi" or browsers="qwap" or browsers="jigs" or browsers="java" or browsers="alca" or browsers="wapj" or browsers="cdr/" or browsers="fetc" or browsers="r380" or browsers="wind" or  browsers="mozi" or browsers="m3ga" or browsers="robo" or browsers="http" or browsers="iask" or browsers="max-" then
if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else 
If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if
response.end 
end if
%>


<%
Dim IpAddr,Values2
Value2=Request.ServerVariables("REMOTE_ADDR")'原始客户端IP地址
'====判断是否移动IP======
IpAddr=Clng(Left(Trim(Replace(Value2,".","")),6))
If (IpAddr<>211103 And IpAddr<211136) Or (IpAddr>211143 And IpAddr<218200) Or (IpAddr>218207 And IpAddr<221130) Or (IpAddr>221131) Then
if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else 
If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if
response.end
end if
%>


<% '判断IP是否已访问该广告
Set rst= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ip where ggid='"&rs("ggid")&"' and ip='"&Value2&"' and time=#"&date()&"#"
rst.Open sql,conn,1,2
if not rs.eof then
'记录已存在，给此用户加pv访问量
Set rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj where ggid='"&rs("ggid")&"' and username='"&rs("username")&" and time=#"&date()&"#'"
rsn.Open sql,conn,1,2
if not rsn.eof then
rsn("zrdjpv")=cint(rsn("zrdjpv")+1)
rsn.UPdate
rsn.close
set rsn=nothing
end if
if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else 
If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if
Response.End
end if 
rst.close
set rst=nothing
%>


<%if Request("ips")="" then%>
<%
Function RndCode()
CodeSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
AmountSet = 62
Randomize
Dim vCode(16), vCodes
For i = 0 To 16
  vCode(i) = Int(Rnd * AmountSet)
  vCodes = vCodes & Mid(CodeSet, vCode(i) + 1, 1)
Next
RndCode=vCodes
End Function
dim sessionid
sessionid=RndCode()

Set rsz= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ip"
rsz.Open sql,conn,1,2
rsz.AddNew
rsz("username")=rs("username")
rsz("ip")=Request.ServerVariables("REMOTE_ADDR")
rsz("ips")=sessionid&Request("c")
rsz("ggid")=rs("ggid")
rsz.Update
rsz.close
set rsz=nothing
rs.close
set rs=nothing
rss.close
set rss=nothing
conn.close
set conn=nothing
Response.write "<card id='card1' title='正在进入..' ontimer='c.asp?c="&Request("c")&"&amp;ips="&sessionid&Request("c")&"'><timer value='10'/>"
%>
</card>
</wml>
<%end if%>


<% 
if Request("ips")<>"" then
Set rsz= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ip where username='"&rs("username")&"' and ggid='"&rs("ggid")&"' and ips='"&Request("ips")&"' and time=#"&date()&"# and zt=0"
rsz.Open sql,conn,1,2
if not rsz.eof then
Set rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj where ggid='"&rs("ggid")&"' and username='"&rs("username")&" and time=#"&date()&"#'"
rsn.Open sql,conn,1,2
'记录存在修改加访问IP,访问pv数量
if not rsn.eof then
rsn("zrdjip")=cint(rsn("zrdjip")+1)
rsn("zrdjpv")=cint(rsn("zrdjpv")+1)
rsn.UPdate
rsn.close
set rsn=nothing
else
'记录不存在新增访问IP,访问pv数量
et rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj"
rsn.Open sql,conn,1,2
rsn.addnew
rsn("username")=rs("username")
rsn("ggid")=rs("ggid")
rsn("zrdjip")=cint(rsn("zrdjip")+1)
rsn("zrdjpv")=cint(rsn("zrdjpv")+1)
rsn.Update
rsn.close
set rsn=nothing
end if
end if
'开始转到广告页面地址
if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else 
If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if

rs.close
set rs=nothing
rss.close
set rss=nothing
conn.close
set conn=nothing
end if
%>