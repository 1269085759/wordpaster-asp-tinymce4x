<%
'ʹ�÷���
'   Dim pathTool
'   Set pathTool = New CPathTool
'   pathTool.MakeDir("D:\\Soft\\System")
'   Set pathTool = Nothing

Function EndsWith(a,b)
    Dim lastV : lastV = Mid(a,Len(a),1)
    Dim cv : cv = StrComp(lastV,b)
    If cv = 0 Then
        EndsWith = True
    Else 
        EndsWith = False
    End If
End Function

'
Function PathCombin(a,b)
    If EndsWith(a,Chr(92)) Then
        PathCombin = a & b
    Else
        PathCombin = a & "\\" & b
    End If
End Function

Sub Println(v)
    Response.Write v
    Response.Write "<br/>"
End Sub

'p=d:\\soft\\safe\\
'p=d:\\soft\\safe
' MakePathLoc("D:\\Soft\\Safe")
Sub MakePathLoc(ByVal p)    
    'ȥ�������\\    
    '��\\��β
    If EndsWith(p,Chr(92)) Then    
        p = Mid(p,0,Len(p))
    End If

    Dim fso,path
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    Dim fds : fds = Split(p,Chr(92))
    Dim root : root = fds(0)    
    For i = 1 To UBound(fds)
        root = PathCombin( root,fds(i) )
        If Not fso.FolderExists(root) Then fso.CreateFolder(root)
    Next
    
    Set fso = Nothing
End Sub

'�Զ�ת�����·��
'MakePathSvr("../upload/")
Sub MakePathSvr(p)
    p = Server.MapPath(p)
    'ȥ�������\\    
    '��\\��β
    If EndsWith(p,Chr(92)) Then    
        p = Mid(p,0,Len(p))
    End If

    Dim fso,path
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    Dim fds : fds = Split(p,Chr(92))
    Dim root : root = fds(0)
    Dim i : i = 1
    For i = 1 To UBound(fds)
        root = PathCombin( root,fds(i) )        
        If Not fso.FolderExists(root) Then fso.CreateFolder(root)
    Next
    
    Set fso = Nothing        
End Sub 
%>