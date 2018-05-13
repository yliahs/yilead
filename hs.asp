
<%
'=============函数===================

'┏━━━━━━━━━━━━━━━━━━━┓
'┃财务进出帐				       ┃
'┗━━━━━━━━━━━━━━━━━━━┛
'参数说明
'username  用户帐号
'jine      金额
'lx        类型（0为减,1为加）
'jysm      交易日志说明

sub money(username,jine,lx,jysm)
if username<>"" and jine<>"" then
'---------修改用户金额
exec="select * from username where username='"&username&"'"
		set rsaa=server.createobject("adodb.recordset")
		rsaa.open exec,conn,1,3
		if lx=1 then
				rsaa("money")=rsaa("money")+jine
			rsaa.update
	
		else
		if rsaa("money")>=jine then
		rsaa("money")=rsaa("money")-jine
		rsaa.update
			else
			zt=0
				money="错误：帐号余额不足！"
				Response.End
			end if
			end if
			rsaa.close
		set rsaa=nothing
end if

'=====================财务写入日志

exec="select * from cwrz"
		set rsab=server.createobject("adodb.recordset")
		rsab.open exec,conn,1,2
		rsab.addnew
		rsab("username")=username
		rsab("money")=jine
		rsab("sm")=jysm
		rsab.update
		rsab.close
		set rsab=nothing

end sub


'┏━━━━━━━━━━━━━━━━━━━┓
'┃检测远程站点状态				 ┃
'┗━━━━━━━━━━━━━━━━━━━┛
'参数说明
'urlcs   要检测的URL
'csurl   返回值
function csurl(urlcs)
	set XMLHTTP = Server.CreateObject("Microsoft.XMLHTTP") 
	XMLHTTP.open "POST",urlcs,false 
	'XMLHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	XMLHTTP.send()
	csurl=XMLHTTP.status
	set XMLHTTP = nothing 
end function


'┏━━━━━━━━━━━━━━━━━━━┓
'┃检测远程站点是否包含非法信息		┃
'┗━━━━━━━━━━━━━━━━━━━┛
'返回值   如果有非法信息返回非法关健字，如果没有返回False
function xxurl(Url)
dim GetXmlHttp
set GetXmlHttp=server.Createobject("Microsoft.XMLHTTP")
GetXmlHttp.open "Get",url,false,"",""
GetXmlHttp.Send 
GetXmlText=GetXmlHttp.Responsetext
set GetXmlHttp=Nothing
dim FilterWord
FilterWord = "六合彩|赌博|反动|暴力|做爱|性爱|性交|性事|房事|A片|日逼|麻逼|操逼|B穴|逼穴|高潮|性欲|色欲|情欲|情色|X夜|失身|春宵|轮奸|强奸|绳虐|叫床|赤裸|裸体|裸奔|波波|两性狂情|禁片|A级|三级|热辣|风骚|骚逼|乳房|巨波|调情|初夜|好色|失身|成人|春潮|处女|脱衣|禽兽|手淫|自慰|淫荡"'过滤名称

if GetXmlText<>"" then
dim text,j
		text = Split(FilterWord,"|")
		For j = 0 to Ubound(text)
		If Instr(GetXmlText,text(j))>0 Then
			xxurl ="关键字："&text(j)&""	
			Exit Function
       end if
       Next
	   end if
	   xxurl = False
end function


'┏━━━━━━━━━━━━━━━━━━━┓
'┃检测站长广告连接广告语	         ┃
'┗━━━━━━━━━━━━━━━━━━━┛
'参数说明
'url      要检测的URL
'userid   用户ID
'ggid     广告ID
'titeurl  返回值

Function titeurl(url,userid,ggid)
dim gettite
set gettite=server.Createobject("Microsoft.XMLHTTP")
gettite.open "Get",url,false,"",""
gettite.Send 
titeurll=gettite.Responsetext
set gettite=Nothing
uu=Request.ServerVariables("SCRIPT_NAME")
uu=Replace(uu,"http://","")
uu=split(trim(uu),"/")
aaurl="http://"&uu(0)&"/c?c="&userid&"@"&ggid&""
start=Instr(titeurll,aaurl)
over =Instr(titeurll,"</wml>")
body=mid(titeurll,start,over-start)
body=split(trim(body),"</a>")
bod=Replace(body(0),"<a href='","")
bod=Replace(body(0),"<a href=","")
bod=Replace(body(0),aaurl,"")
bod=Replace(body(0),"'>","")
bod=Replace(body(0),">","")
titeurl = bod
End Function

Function XmlPostt(url,usee)
	Dim xml
	Set xml = Server.CreateObject("Microsoft.XMLHTTP")
	xml.Open "get",Url,False
	xml.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
	xml.Send(usee)
	XmlPostt = Xml.ResponseText
	Set xml = Nothing
End Function

%>