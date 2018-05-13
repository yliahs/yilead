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
<!--#INCLUDE file="../CF_Sql.asp"-->
<!--#INCLUDE file="../hs.asp"-->
<!--#INCLUDE file="../hx.asp"-->
<!--#INCLUDE file="css.css"-->
<!--#INCLUDE file="../date.asp"-->
<%if session("wztite")="" or session("img")="" or session("description")="" or session("keywords")="" or session("gdgg")="" or session("url")="" or session("dl")="" then
set objgbrs=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where id=1"
  objgbrs.open sql,conn
  dim wztite,img,description,keywords
  session("wztite")=objgbrs("title")
  session("img")=objgbrs("img")
  session("description")=objgbrs("description")
  session("keywords")=objgbrs("keywords")
  session("gdgg")=objgbrs("gdgg") 
  session("dl")=objgbrs("dl")
  session("url")=objgbrs("url")
  objgbrs.close
  set objgbrs=nothing
  end if %>
<%
set objgbrs=Server.CreateObject("ADODB.Recordset")
  sql="select * from admin where id=1"
  objgbrs.open sql,conn
  if objgbrs("name")<>Request.Cookies("admin") or objgbrs("pass")<>Request.Cookies("pass") then
Response.Redirect "login.asp"
 Response.end
 end if
  objgbrs.close
  set objgbrs=nothing
%>