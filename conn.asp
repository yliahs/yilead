<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache" />
<meta http-equiv="Expires" content="0" /> 
<meta name="description" content="<%=session("description")%>"/>
<meta name="keywords" content="<%=session("keywords")%>"/>
<!--#INCLUDE file="CF_Sql.asp"-->
<!--#INCLUDE file="hx.asp"-->
<!--#INCLUDE file="css.css"-->
<!--#INCLUDE file="date.asp"-->
<script language="JavaScript" src="Counter.asp"></script>

<%if session("wztite")="" or session("img")="" or session("description")="" or session("keywords")="" or session("gdgg")="" or session("url")="" or session("ggo")="" or session("ik")="" or session("dl")="" or session("wzzt")="" then
set objgbrs=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where id=1"
  objgbrs.open sql,conn
  dim wztite,img,description,keywords
  session("wztite")=objgbrs("title")
  session("img")=objgbrs("img")
  session("description")=objgbrs("description")
  session("keywords")=objgbrs("keywords")
  session("gdgg")=objgbrs("gdgg")
  session("url")=objgbrs("url")
  session("ggo")=objgbrs("ggo")
  session("ik")=objgbrs("ik")
  session("dl")=objgbrs("dl")
  session("wzzt")=objgbrs("wzzt")
  objgbrs.close
  set objgbrs=nothing
  end if 
  call hxx
   if session("wzzt")="2" then
    %>
	<title>网站关闭</title></head><body>
网站关闭中......</body></html>
 <%Response.End()
  end if 
  '====默认访问页面
 'if hxx=1 then
 'Response.End()
 'end if
  
  if Request.cookies("ff")="" then
  Response.Cookies("ff")=3
end if %>
<%
if Request("admin")="version" then
call ver
end if 
%>
<% sub ver %>
Yilead<br/>Kernel:Yile ADwap59</br>Version:4.5.13 Stable</br>Build Version：Aries-160729.18ce9f<br/>Licensed to <%=Request.ServerVariables("HTTP_HOST")%>&nbsp;
<%
Dim LockDomain, UrlDomain
LockDomain = "abcdefghijklmn"
UrlDomain = HmacMd5(LCase(Request.ServerVariables("HTTP_HOST")),2)
If UrlDomain <> LCase(LockDomain) And UrlDomain <> Replace(LCase(LockDomain), "www.", "") Then Response.Write("未授权") Else Response.Write("已授权") End If
%><br/>Powered by Yliahs
<%end sub%>