<!--#INCLUDE file="conn.asp"-->
<title>广告列表</title>	
</head>
<body>
    <div id="all">

<p><div class="p7" align="center"><img src="<%=session("img")%>" width="220" height="40" /></div></p>
<!--#INCLUDE file="top.asp"-->
<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

      <div class="p12"> <img src="images/lalala.gif" width="14" height="14" />广告代码</div>
	 
	 
<%
	  set rs=Server.CreateObject("ADODB.Recordset")
  sql="SELECT * From gglb Order By px asc"
  rs.open sql,conn
  Response.write "<p>"
  do while ((not rs.EOF))  
  dda=rs("id")  
  if cint(Request("action"))=cint(rs("id")) then
  Response.write "<span class='tab'><span>"&rs("title")&"</span></span>&nbsp;"
  else
   Response.write "<a href='gg.asp?action="&rs("id")&"'>"&rs("title")&"</a>&nbsp;"
  end if
   rs.MoveNext
loop 
Response.write "</p>"
rs.close
set rs=nothing 
  %>
	 
  <%
  set rs=Server.CreateObject("ADODB.Recordset")
  if Request("action")="" then
  sql="SELECT * From ad where gglb='"&dda&"' and ggzt=1 Order By id desc"
  else
  sql="SELECT * From ad where gglb='"&Request("action")&"' and ggzt=1 Order By id desc"
  end if
  rs.open sql,conn
  dim i
  i=0
   do while ((not rs.EOF))  
     i=i+1
  %>

   <div style="border:1px inset #CCC; width:100%; text-align:center; height:44px; " >
       <div>
            <div style="width:25%; height:22px; float:left">
                 <span id="Repeater1_ctl00_Label1" style="color:Blue;"><font color="#333333">广告<%=i%></font></span> 
            </div>
             <div style="width:75%; height:22px; float:left">
                 <a id="Repeater1_ctl00_HyperLink1" href="gg_list.asp?id=<%=rs("id")%>"><%=rs("title")%></a></span> 
            </div>
        </div>
          <div>
            <div style="width:25%; height:22px; float:left">
                   <a id="Repeater1_ctl00_HyperLink_Sq" href="gg_list.asp?id=<%=rs("id")%>"><font color="#0000FF">获取代码</font></a>  </span> 
            </div>
             <div style="width:75%; height:22px; float:left">
                价格:&nbsp;<span id="Repeater1_ctl00_Label4" style="color:#000F00;"><%=rs("money")%>元/<%=rs("jmoney")%></span></span> 
            </div>
        </div>
   </div>
      
 <%   
 rs.MoveNext
loop 
rs.close
set rs=nothing 

%>      
<p class="px"><img src="images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="db.asp"-->
</div>
</body>
</html>
