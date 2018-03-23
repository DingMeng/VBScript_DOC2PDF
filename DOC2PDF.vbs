Dim fso,fld,Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Path = fso.GetParentFolderName(WScript.ScriptFullName) '获取脚本所在文件夹字符串
Set fld=fso.GetFolder(Path) '通过路径字符串获取文件夹对象

Dim Sum,IsChooseDelete,ThisTime
Sum = 0
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileList.txt",2,true)

Call LogOut("开始遍历文件")
Call TreatSubFolder(fld) '调用该过程进行递归遍历该文件夹对象下的所有文件对象及子文件夹对象

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '遍历该文件夹对象下的所有文件对象
        If UCase(fso.GetExtensionName(File)) ="DOC" or UCase(fso.GetExtensionName(File)) ="DOCX" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders '递归遍历子文件夹对象
        TreatSubFolder subfld
    Next
End Sub
List.close

Call LogOut("文件遍历已完成，已找到" & Sum & "个word文档")


'创建Word对象，兼容WPS
Const wdFormatPDF = 17
On Error Resume Next
Set WordApp = CreateObject("Word.Application")
' try to connect to wps
If WordApp Is Nothing Then '兼容WPS
    Set WordApp = CreateObject("WPS.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("KWPS.Application")
        If WordApp Is Nothing Then
            LogOut("未检测到office2010及以上的版本，转换失败！")
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

WordApp.Visible=false '设置视图不可见

Sum = 0
Dim FilePath,FileLine
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine=List.ReadLine
    If FileLine <> "" and Mid(FileLine,1,2) <> "~$" Then
        Sum = Sum + 1 '获取用户修改后的文件列表行数
    End If
loop
List.close
Dim Finished
Finished = 0
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    If Mid(FilePath,1,2) <> "~$" Then '不处理word临时文件
        Set objDoc = WordApp.Documents.Open(FilePath)
        'WordApp.Visible=false '设置视图不可见（避免运行时因为各种问题导致的可见）
        '上面这行有问题，现在遇到大批量有啥宏定义的运行起来就是一闪一闪的，还不如没有
        If WordApp.Visible = true Then
            WordApp.ActiveDocument.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        objDoc.SaveAs Left(FilePath,InstrRev(FilePath,".")) & "pdf", wdFormatPDF '另存为PDF文档
        LogOut("文档" & FilePath & "已转换完成。(" & Finished & "/" & Sum & ")")
        WordApp.ActiveDocument.Close  
        Finished = Finished + 1
    End If
    If IsChooseDelete = vbYes Then
        fso.deleteFile FilePath
        LogOut("文件" & FilePath & "已被成功删除")
    End If
loop
'扫尾处理开始
List.close
LogOut("文档转换已完成")
LogFile.close 
'ConvertFileList.txt和log.txt要自动删除的请去掉下面两行开头单引号
'fso.deleteFile "ConvertFileList.txt"
'fso.deleteFile "log.txt"

Dim Msg
Msg = "已成功转换" & Finished & "个文件"
If IsChooseDelete = vbYes Then
    Msg=Msg + "并成功删除源文件"
End If
Set fso = nothing
WordApp.Quit
Wscript.Quit
