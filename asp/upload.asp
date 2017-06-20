<% @Language=vbscript Codepage=936 %>
<%
Option Explicit
Response.Buffer=True
%>
<!--#include file="UpLoadClass.asp"-->
<!--#include file="PathTool.asp"-->
<%
'更新记录：
'	2012-3-28 增加创建年月
dim lngUpSize,uploader,intError
	Set uploader = new UpLoadClass
	uploader.TotalSize= 10485760'10MB
	uploader.MaxSize  = 10000*1024
	uploader.FileType = "gif/jpg/png/bmp"
	'服务器域名
	Dim domain : domain = "http://localhost:90/asp/"	
	Dim dateNow : dateNow = Date()
	'存储路径：upload/年/月/日/
	Dim filePath : filePath = "upload/" & Year(dateNow) & "/" & Month(dateNow) & "/" & Day(dateNow) & "/"
	uploader.Savepath = filePath
    '自动创建目录
    MakePathSvr(filePath)	

	lngUpSize = uploader.Open()
	intError = uploader.Form("photo2_Err")
	'输出文件名称和路径：2011-09-10-5-52-255252.jpg'
	response.Write(domain & filePath & uploader.Form("ServerFileName"))
	if lngUpSize>uploader.MaxSize then
%>
		<script language="javascript">
		<!--
			alert("您上传的文件最大不能超过10M!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	if intError=-1 then
%>
		<script language="javascript">
		<!--
			alert("您没有上传任何文件，请重新上传!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	Set uploader = nothing
%>
