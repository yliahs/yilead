<% Response.ContentType="text/vnd.wap.wml; charset=utf-8" %>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="max-age=0"/>
<meta http-equiv="Cache-Control" content="no-cache"/>
<meta http-equiv="Expires" content="Mon, 1 Jan 1990 00:00:00 GMT"/>
</head>
<!--#INCLUDE file="../date.asp"-->

  <%
  size=trim(Request("c"))
  myarr=split(size,"@")
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&trim(myarr(0))&"' and ggid="&trim(myarr(1))&""
  rs.open sql,conn,1,2
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

 <%Function ggurl
 set rss=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where id="&rs("ggid")&""
  rss.open sql,conn
 if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else
 If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if
End Function%>
  
  <%'判断来源地址
  set rsaa=Server.CreateObject("ADODB.Recordset")
  sql="select fwxz from admin where id=1"
  rsaa.open sql,conn
  if rsaa("fwxz")=1 then
  dim urls
  urls=request.ServerVariables("HTTP_REFERER")
	'if urls="" then
'Response.write "<card id='card1' title='访问错误'>"
 'Response.write "来源址址：空<br/>来源地址为空不能访问广告,请将广告代码挂到您已申请的站点里面，谢谢合作！<br/>" 
' Response.write "</card></wml>"
 'Response.end
'end if 
	urls=Replace(urls,"http://","")
	urls=split(trim(urls),"/")
  size=trim(Request("c"))
  myarr=split(size,"@")
  set rst=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&trim(myarr(0))&"' and ggid="&trim(myarr(1))&" and url='"&trim(urls(0))&"' and zt=2"
  rst.open sql,conn

if rst.eof then
  Response.write "<card id='card1' title='访问错误'>"
 Response.write "来源站点："&urls(0)&"<br/>此来源站点没有申请此广告或没有审核通过！<br/>" 
 Response.write "</card></wml>"
 Response.end
 end if
rst.close
set rst=nothing	 
end if
rsaa.close
set rsaa=nothing
  %>
<%Function ua
'======屏蔽UA客户端=========
browsers=Lcase(Left(Request.ServerVariables("HTTP_USER_AGENT"),4))
if browsers="oper" or browsers="winw" or browsers="wap/" or browsers="wapi" or browsers="mc21" or browsers="up.b" or browsers="upg1" or browsers="upsi" or browsers="qwap" or browsers="jigs" or browsers="java" or browsers="alca" or browsers="wapj" or browsers="cdr/" or browsers="fetc" or browsers="r380" or browsers="wind" or  browsers="mozi" or browsers="m3ga" or browsers="robo" or browsers="http" or browsers="iask" or browsers="max-" then
call ggurl
response.end 
end if
End Function
'call ua
%>
<% 
Dim IpAddr,Values2
Value2=Request.ServerVariables("REMOTE_ADDR")'原始客户端IP地址
IpAddr=Clng(Left(Trim(Replace(Value2,".","")),6))
'====判断是否移动IP======
Function ipp
If (IpAddr<>211103 And IpAddr<211136) Or (IpAddr>211143 And IpAddr<218200) Or (IpAddr>218207 And IpAddr<221130) Or (IpAddr>221131) Then
call ggurl
response.end
end if
End Function
'call ipp
%>

<% '判断IP是否已访问该广告
Set rst= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ip where ggid="&rs("ggid")&" and ip='"&Value2&"' and time=#"&date&"#"
rst.Open sql,conn,1,2

if not rst.eof then
'记录已存在，给此用户加pv访问量
Set rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj where ggid='"&rs("ggid")&"' and username='"&rs("username")&"' and time=#"&date&"#"
rsn.Open sql,conn,1,2
if not rsn.eof then
rsn("zrdjpv")=cint(rsn("zrdjpv")+1)
rsn.UPdate
rsn.close
set rsn=nothing
end if
Set rsb= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggfw where ggid='"&rs("ggid")&"' and time=#"&date()&"#"
rsb.Open sql,conn,1,2
if not rsb.eof then
rsb("pv")=rsb("pv")+1
rsb.update
else 
rsb.addnew
rsb("ggid")=rs("ggid")
rsb("title")=rs("ggtitle")
rsb("pv")=rsb("pv")+1
rsb.update
end if
set rsb=nothing
call ggurl
Response.End
rst.close
set rst=nothing
end if%>

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
rsz("url")=request.ServerVariables("HTTP_REFERER")
rsz("xg")=0
rsz.Update
rsz.close
set rsz=nothing
rs.close
set rs=nothing
rss.close
set rss=nothing
conn.close
set conn=nothing
Response.cookies("ips")=sessionid&Request("c")
Response.write "<card id='card1' title='正在进入..' ontimer='indexx.asp?c="&Request("c")&"'><timer value='2'/>"
%>
</card>
</wml>
