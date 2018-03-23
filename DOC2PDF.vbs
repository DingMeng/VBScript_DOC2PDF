Dim fso,fld,Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Path = fso.GetParentFolderName(WScript.ScriptFullName) '��ȡ�ű������ļ����ַ���
Set fld=fso.GetFolder(Path) 'ͨ��·���ַ�����ȡ�ļ��ж���

Dim Sum,IsChooseDelete,ThisTime
Sum = 0
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileList.txt",2,true)

Call LogOut("��ʼ�����ļ�")
Call TreatSubFolder(fld) '���øù��̽��еݹ�������ļ��ж����µ������ļ��������ļ��ж���

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '�������ļ��ж����µ������ļ�����
        If UCase(fso.GetExtensionName(File)) ="DOC" or UCase(fso.GetExtensionName(File)) ="DOCX" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders '�ݹ�������ļ��ж���
        TreatSubFolder subfld
    Next
End Sub
List.close

Call LogOut("�ļ���������ɣ����ҵ�" & Sum & "��word�ĵ�")


'����Word���󣬼���WPS
Const wdFormatPDF = 17
On Error Resume Next
Set WordApp = CreateObject("Word.Application")
' try to connect to wps
If WordApp Is Nothing Then '����WPS
    Set WordApp = CreateObject("WPS.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("KWPS.Application")
        If WordApp Is Nothing Then
            LogOut("δ��⵽office2010�����ϵİ汾��ת��ʧ�ܣ�")
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

WordApp.Visible=false '������ͼ���ɼ�

Sum = 0
Dim FilePath,FileLine
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine=List.ReadLine
    If FileLine <> "" and Mid(FileLine,1,2) <> "~$" Then
        Sum = Sum + 1 '��ȡ�û��޸ĺ���ļ��б�����
    End If
loop
List.close
Dim Finished
Finished = 0
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    If Mid(FilePath,1,2) <> "~$" Then '������word��ʱ�ļ�
        Set objDoc = WordApp.Documents.Open(FilePath)
        'WordApp.Visible=false '������ͼ���ɼ�����������ʱ��Ϊ�������⵼�µĿɼ���
        '�������������⣬����������������ɶ�궨���������������һ��һ���ģ�������û��
        If WordApp.Visible = true Then
            WordApp.ActiveDocument.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        objDoc.SaveAs Left(FilePath,InstrRev(FilePath,".")) & "pdf", wdFormatPDF '���ΪPDF�ĵ�
        LogOut("�ĵ�" & FilePath & "��ת����ɡ�(" & Finished & "/" & Sum & ")")
        WordApp.ActiveDocument.Close  
        Finished = Finished + 1
    End If
    If IsChooseDelete = vbYes Then
        fso.deleteFile FilePath
        LogOut("�ļ�" & FilePath & "�ѱ��ɹ�ɾ��")
    End If
loop
'ɨβ����ʼ
List.close
LogOut("�ĵ�ת�������")
LogFile.close 
'ConvertFileList.txt��log.txtҪ�Զ�ɾ������ȥ���������п�ͷ������
'fso.deleteFile "ConvertFileList.txt"
'fso.deleteFile "log.txt"

Dim Msg
Msg = "�ѳɹ�ת��" & Finished & "���ļ�"
If IsChooseDelete = vbYes Then
    Msg=Msg + "���ɹ�ɾ��Դ�ļ�"
End If
Set fso = nothing
WordApp.Quit
Wscript.Quit
