<% @Language=vbscript Codepage=936 %>
<%
Option Explicit
Response.Buffer=True
%>
<!--#include file="UpLoadClass.asp"-->
<!--#include file="PathTool.asp"-->
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
	Dim dateNow : dateNow = Date()
	'�洢·����upload/��/��/��/
	Dim filePath : filePath = "upload/" & Year(dateNow) & "/" & Month(dateNow) & "/" & Day(dateNow) & "/"
	uploader.Savepath = filePath
    '�Զ�����Ŀ¼
    MakePathSvr(filePath)	

	lngUpSize = uploader.Open()
	intError = uploader.Form("photo2_Err")
	'����ļ����ƺ�·����2011-09-10-5-52-255252.jpg'
	response.Write(domain & filePath & uploader.Form("ServerFileName"))
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
