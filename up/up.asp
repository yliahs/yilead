<!--#INCLUDE file="../conn.asp"-->
<!--#include FILE="upload.inc"-->
<title>身份证件上传</title>	
</head>
<body>
    <div id="all">


<%if Request.Cookies("username")="" then%>
<%Response.Redirect "/login.asp"
Response.end%>
<%else%>
<p><div class="p7" align="center"><img src="/<%=session("img")%>" width="220" height="40" /></div></p>
<p><div class="p31">合作ID:<%=Request.Cookies("id")%>|<a id="Top1_HyperLink1" href="/login.asp?action=zx">注销登陆</a></div></p>
<%end if%>

<p><div class="p2"><marquee scrolldelay="110" scrollamount="2"><span id="Top1_Label2"><%=session("gdgg")%></span></marquee></div></p>

<div class="p13">&nbsp;<img src="../images/dada.gif" />身份证件上传</div>	 
<%
action=Request("action")
if action="xg" then
sq1="select * from sfz where username='"&Request.Cookies("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.Recordset")
rs1.open sq1,conn,1,2
if not(rs1.eof and rs1.bof) then
if rs1("zt")=2 then
response.Redirect"err.asp?id=4"
Response.end
end if
end if
end if

If Right(rsformPath,1)<>"\" Then rsformPath=rsformPath&"\"
%>
<%
dim upload,file,formName,formPath,iCount
Dim sjs,fname,i,l,newfilelist,uploaddir,filename,ii
''--------------------------
Server.ScriptTimeOut=999999
set upload=new upload_5xsoft 
''---------------------------获得变量值
fullpath=Server.Mappath("pic")&"\"
''--------------------------保存路径
addip=request.serverVariables("remote_host")
title=upload.form("title")
explain=upload.form("explain")
addid=session("id")


formPath=rsformPath
fullpath=fullpath&rsformPath
On Error Resume Next 
Set fso = CreateObject("Scripting.FileSystemObject")
Set fldr = fso.GetFolder(fullpath)
If err<>0 Then 
response.Redirect"err.asp?id=6"
Response.end
end if
''--------------------------
iCount=0
for each formName in upload.objFile ''列出所有上传了的文件
set file=upload.file(formName) ''生成一个文件对象
size=file.filesize
''-------------------------限制文件大小
rsfilesize="100"
If size>rsfilesize*1024 Then 
response.Redirect"err.asp?id=3"
Response.end
end if
If  size=0  Then 
response.Redirect"err.asp?id=5"
response.end
end if
''-------------------------获得文件类型
filetype=file.filename
i=InStrRev(filetype,".")
l=Len(filetype)
If i>0 Then
  filetype=Right(filetype,l-i+1)
End If
''-------------------------限制文件类型
filetype=LCase(filetype)
filetype=Replace(filetype,".","")
filetype=CStr(filetype)
rsallowedfile="gif,jpg,png,jpeg"
If  InStr(rsallowedfile,filetype)=0  Then 
response.Redirect"err.asp?id=2"
Response.end
end if
''----------------取得新文件名
If rsnamekind=1 Then
fname=file.filename
Else
RANDOMIZE
sjs=INT((99-00+1)*RND+00)
fname=year(date)&month(date)&day(date)&hour(time())&minute(time())&second(time())&sjs
fname=fname&"."&filetype
End if
''-------------------------检验文件是否存在
'If rsforceup="F" then
'sq1="select * from data where filesize='"&size&"'"
'Set Rs1 = Server.CreateObject("Adodb.Recordset")
'rs1.open sq1,conn,1,2
if action<>"xg" then
sq1="select * from sfz where username='"&Request.Cookies("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.Recordset")
rs1.open sq1,conn,1,2
If Not rs1.bof Or Not rs1.eof Then 
response.Redirect"err.asp?id=4&size="&size&""
Response.end
Else
sq1="select * from sfz"
Set Rs1 = Server.CreateObject("Adodb.Recordset")
rs1.open sq1,conn,1,2
End If
else
sq1="select * from sfz"
Set Rs1 = Server.CreateObject("Adodb.Recordset")
rs1.open sq1,conn,1,2
end if
''--------------------------
if file.filesize>0 then         ''如果 FileSize > 0 说明有文件数据
file.SaveAs fullpath&fname ''保存文件
iCount=iCount+1
msg="文件上传成功！"
End  If
datadir=formPath&fname
addtime=now()
if action<>"xg" then
rs1.addnew
end if
rs1("url")=fname
rs1("timee")=addtime
rs1("username")=Request.Cookies("username")
rs1("zt")=1 '1为未审核，2为已审核，3为未通过审核
rs1.update
rs1.close
set file=nothing
next
set upload=nothing  '删除此对象
%>
<card  title="上传成功"><p>
<%
'size=CStr(size)
'sq2="select * from data where filesize='"&size&"' order by id desc"
'Set Rs2 = Server.CreateObject("Adodb.Recordset")
'rs2.open sq2,conn,1,2
'id=rs2("id")
'title=rs2("title")
'explain=rs2("explain")
'size=rs2("filesize")
'addtime=rs2("addtime")
'rs2.close
'Set rs2=nothing
%>
<p>
证件上传成功！请等待或联系客服审核！<br/>
-------------<br/>
<%'路径：""&fname&""<br/>%>
大小:<%size=Round(size/1024,2)
If Left(size,1)="." Then size="0"&size
%>
<%=size%>KB<br/>
上传时间:<%=addtime%><br/>
<img src="/up/pic/<%=fname%>" width="150" height="120"/>
</p>

-------------<br/>
<p class="px"><img src="/images/fanhui.gif" width="16" height="9" alt='返回首页'/><a href="/">返回首页</a></p>
<!--#INCLUDE file="../db.asp"-->
</div>
</body>
</html>