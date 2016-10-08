<% @Language=vbscript Codepage=936 %>
<%
Option Explicit
Response.Buffer=True
%>
<!--#include file="UpLoadClass.asp"-->
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
	Dim pathRoot : pathRoot = "upload/"
	Dim dateNow : dateNow = Date()
	'2012-3-7
	Dim timeCur : timeCur = Year(dateNow) & Month(dateNow) & "/" & Day(dateNow) & "/"
	uploader.Savepath = pathRoot & timeCur
	
	'自动创建上传文件夹
	Dim folderRoot : folderRoot = Server.MapPath(pathRoot)
	Dim folderYM : folderYM = Server.MapPath(pathRoot & Year(dateNow) & Month(dateNow))
	Dim folder : folder = server.MapPath(uploader.Savepath)
	Dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	'创建根目录
	If(not fs.FolderExists(folderRoot)) Then
		fs.CreateFolder(folderRoot)
	End If
	'创建年月目录
	If(Not fs.FolderExists(folderYM)) Then
		fs.CreateFolder(folderYM)
	End If
	'创建子目录
	If(not fs.FolderExists(folder)) Then
		fs.CreateFolder(folder)
	End If

	lngUpSize = uploader.Open()
	intError = uploader.Form("photo2_Err")
	'输出文件名称和路径：2011-09-10-5-52-255252.jpg'
	Response.Write(domain & uploader.Savepath & uploader.Form("ServerFileName"))
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
