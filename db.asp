<%conn.close
set conn=nothing
   'Dim ScriptAddresss,Servernames,qss
   ScriptAddresss = CStr(Request.ServerVariables("SCRIPT_NAME"))
   Servernames = CStr(Request.ServerVariables("Server_Name"))
   qss=Request.QueryString
   qss=Replace(qss,"&","&amp;")
%>
<div class="p6">
   <table width="100%" height="50px" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td align="center"><%=session("dl")%><br/>Powered By yile</td>      
  </tr>
  </table>
  </div>