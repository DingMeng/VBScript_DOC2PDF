Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Dim fld ' as object
dim Path ' As string
Path = fso.GetParentFolderName(WScript.ScriptFullName) '��ȡ�ű������ļ����ַ���
Set fld=fso.GetFolder(Path) 'ͨ��·���ַ�����ȡ�ļ��ж���

Dim Sum,IsChooseDelete,ThisTime
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
        If fso.GetExtensionName(File) ="doc" or fso.GetExtensionName(File)="docx" Then
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

If MsgBox("�ļ���������ɣ����ҵ�" & Sum & "��word�ĵ�����ϸ�б���" & vbCrlf & "ConvertFileList.txt" & vbCrlf & "�Ƿ���Щ�ĵ�ת��ΪPDF��", vbYesNo + vbInformation, "�ĵ��������") = vbYes Then
    If MsgBox("�Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo+vbInformation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?") = vbYes Then
        IsChooseDelete = MsgBox("���ٴ�ȷ�ϣ��Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo + vbExclamation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?")
    End If
else
    Msgbox("��ȡ��ת������")
    Wscript.Quit
End If
MsgBox "���ڿ�ʼת��ǰ�˳�����Word�ĵ������ĵ�ռ�ô�����", vbOKOnly + vbExclamation, "����"


Const wdFormatPDF = 17
Set wdapp = CreateObject("Word.Application")'����Word����
wdapp.Visible=false '������ͼ���ɼ�

Dim Finished
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    Set objDoc = wdapp.Documents.Open(FilePath)
    objDoc.SaveAs Left(FilePath,InstrRev(FilePath,".")) & "pdf", wdFormatPDF '���ΪPDF�ĵ�
    LogOut("�ĵ�" & FilePath & "��ת����ɡ�(" & Finished & "/" & Sum & ")")
    wdapp.ActiveDocument.Close  
    Finished = Finished + 1
    If IsChooseDelete = vbYes Then
        fso.deleteFile FilePath
        LogOut("�ļ�" & FilePath & "�ѱ��ɹ�ɾ��")
    End If
loop
'ɨβ����ConvertFileList.txt��log.txtҪ�Զ�ɾ������ȥ���������п�ͷ������
'fso.deleteFile "ConvertFileList.txt"
'fso.deleteFile "log.txt"
List.close
LogOut("�ĵ�ת�������")
LogFile.close 
Set fso = nothing

Dim Msg
Msg = "�ѳɹ�ת��" & Finished & "���ļ�"
If IsChooseDelete = vbYes Then
    Msg=Msg + "���ɹ�ɾ��Դ�ļ�"
End If
MsgBox Msg
wdapp.Quit
Wscript.Quit
