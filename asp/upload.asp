<% @Language=vbscript Codepage=936 %>
<%
Option Explicit
Response.Buffer=True
%>
<!--#include file="UpLoadClass.asp"-->
<%
'���¼�¼��
'	2012-3-28 ���Ӵ�������
dim lngUpSize,uploader,intError
	Set uploader = new UpLoadClass
	uploader.TotalSize= 10485760'10MB
	uploader.MaxSize  = 10000*1024
	uploader.FileType = "gif/jpg/png/bmp"
	'����������
	Dim domain : domain = "http://localhost:90/asp/"
	Dim pathRoot : pathRoot = "upload/"
	Dim dateNow : dateNow = Date()
	'2012-3-7
	Dim timeCur : timeCur = Year(dateNow) & Month(dateNow) & "/" & Day(dateNow) & "/"
	uploader.Savepath = pathRoot & timeCur
	
	'�Զ������ϴ��ļ���
	Dim folderRoot : folderRoot = Server.MapPath(pathRoot)
	Dim folderYM : folderYM = Server.MapPath(pathRoot & Year(dateNow) & Month(dateNow))
	Dim folder : folder = server.MapPath(uploader.Savepath)
	Dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	'������Ŀ¼
	If(not fs.FolderExists(folderRoot)) Then
		fs.CreateFolder(folderRoot)
	End If
	'��������Ŀ¼
	If(Not fs.FolderExists(folderYM)) Then
		fs.CreateFolder(folderYM)
	End If
	'������Ŀ¼
	If(not fs.FolderExists(folder)) Then
		fs.CreateFolder(folder)
	End If

	lngUpSize = uploader.Open()
	intError = uploader.Form("photo2_Err")
	'����ļ����ƺ�·����2011-09-10-5-52-255252.jpg'
	Response.Write(domain & uploader.Savepath & uploader.Form("ServerFileName"))
	if lngUpSize>uploader.MaxSize then
%>
		<script language="javascript">
		<!--
			alert("���ϴ����ļ�����ܳ���10M!!");
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
			alert("��û���ϴ��κ��ļ����������ϴ�!!");
			history.back();
		//-->
		</script>
<%
		response.end
	end if
	Set uploader = nothing
%>
