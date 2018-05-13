<!--#INCLUDE file="../date.asp"-->
<%
  size=trim(Request("c"))
  myarr=split(size,"@")
  set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from ggurl where userid='"&trim(myarr(0))&"' and ggid="&trim(myarr(1))&""
  rs.open sql,conn %>
  
     <%
set rss=Server.CreateObject("ADODB.Recordset")
  sql="select * from ad where id="&rs("ggid")&""
  rss.open sql,conn
  %>
 <%sub ggurl
 if rs("ggurl")<>"" then
response.Redirect rs("ggurl")
else
 If Instr(rss("url"),"?") Then
response.Redirect ""&rss("url")&"&uid="&rs("userid")&"&gid="&rs("ggid")&""
else
response.Redirect ""&rss("url")&"?uid="&rs("userid")&"&gid="&rs("ggid")&""
end if
end if
 end sub%>
 
<%
dim ztt
ztt=0
Set rsz= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ip where username='"&rs("username")&"' and ggid="&rs("ggid")&" and ips='"&Request.Cookies("ips")&"'  and zt="&ztt&"  and time=#"&date()&"#"
rsz.Open sql,conn,1,2

if not rsz.eof then
rsz("zt")=1
rsz.update
Set rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj where ggid='"&rs("ggid")&"' and username='"&rs("username")&"' and time=#"&date()&"#"
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
set rsn= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggsj"
rsn.Open sql,conn,1,2
rsn.addnew
rsn("username")=rs("username")
rsn("ggid")=rs("ggid")
rsn("gglx")=rs("gglx")
rsn("ggtitle")=rs("ggtitle")
rsn("zrdjip")=cint(rsn("zrdjip")+1)
rsn("zrdjpv")=cint(rsn("zrdjpv")+1)
rsn.Update
rsn.close
set rsn=nothing
end if
end if

Set rsb= Server.CreateObject("ADODB.Recordset")
sql="SELECT * From ggfw where ggid='"&rs("ggid")&"' and time=#"&date()&"#"
rsb.Open sql,conn,1,2
if not rsb.eof then
rsb("ip")=rsb("ip")+1
rsb("pv")=rsb("pv")+1
rsb.update
else
rsb.addnew
rsb("ggid")=rs("ggid")
rsb("title")=rs("ggtitle")
rsb("ip")=rsb("ip")+1
rsb("pv")=rsb("pv")+1
rsb.update
end if
rsb.close
set rsb=nothing
'开始转到广告页面地址
call ggurl

rs.close
set rs=nothing
rss.close
set rss=nothing
conn.close
set conn=nothing
%>