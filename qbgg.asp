<!--#INCLUDE file="conn.asp"-->

<title><%=session("wztite")%></title>
</head>
<body>
<div id="all">   
<div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div>
<div class="p2"><marquee scrolldelay="110" scrollamount="2"><p><span id="Top1_Label2"><%=session("gdgg")%></span></p></marquee></div>
<div class="p13">&nbsp;<img src="images/dada.gif" />全部公告</div><p>
<%
set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT top 99999 * From wzgg Order By time desc"
  rs.open sql,conn
   do while ((not rs.EOF))   
  Response.write "<a id='Repeater1_ctl04_HyperLink1' href='wzgg.Asp?id="&rs("id")&"'>"&rs("title")&"</a><br/>"
  rs.MoveNext
loop 
rs.close
set rs=nothing 

'sfile="/CF_Sql.asp"
'call fsofiledatemofei1(sfile,3074)
'sfile="/conn.asp"
'call fsofiledatemofei1(sfile,1787)
'sfile="/hs.asp"
'call fsofiledatemofei1(sfile,4082)  
%></p>
 <p class="px"><img src="images/fanhui.gif" width="16" height="9"  alt='首页'/><a href="/">返回首页</a></p>
 <!--#INCLUDE file="db.asp"--> 
</div>
</body>
</html>
