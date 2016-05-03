Attribute VB_Name = "ModMain"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>

        Dim fso As Scripting.FileSystemObject

100     Set fso = New Scripting.FileSystemObject

102     With fso

104         If Not .FolderExists(App.Path & "\log") Then

106             .CreateFolder App.Path & "\log"

            End If

        End With

108     Set fso = Nothing

110     MainForm.Show

        '<EhFooter>
        Exit Sub

Main_Err:
        WriteLog Err.Description & vbCrLf & _
           "in RegExpTest.ModMain.Main " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Function ConvertHTML(ByVal Content As String, Optional ByVal intSize As Integer = 0)
    Content = restoreCRLF(Content)
    Content = Replace(Content, vbTab, " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&nbsp;", " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "'", "`", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&lt;", "<", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&gt;", ">", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(10), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(9), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(13), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<BR>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br />", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf, 1, -1, vbBinaryCompare)

    '以上的顺序是有讲究的，不可乱动！
    Dim i As Integer

    For i = 0 To 4

        Content = Replace(Content, "  ", "", 1, -1, vbBinaryCompare)

    Next

    Dim regTmp As VBScript_RegExp_55.RegExp

    Set regTmp = New VBScript_RegExp_55.RegExp
    regTmp.Global = True
    regTmp.MultiLine = True
    regTmp.IgnoreCase = True
    '======================= add by brightdog 去除页面中的干扰码
    regTmp.Pattern = "(<span[^>]*?display\s*?:\s*?none[^>]*?>[\w\W]*?<\/span>)"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<font([^>]+)(0px|0pt)+([^>]*)>([\w\W]*?)<\/font>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<span[^>]*?font\s*?-\s*?size\s*?:\s*(0px|0pt)[^>]*?>([\w\W]*?)<\/span>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<script[^>]*?>([\w\W]*?)<\/script>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    '=======================
    regTmp.Pattern = "(width\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(height\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<em>.*?</em>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    
    regTmp.Pattern = "(<.*?[^>]>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<.*?>)"
    Content = regTmp.Replace(Content, "")
    Content = Trim$(Content)
    
    Set regTmp = Nothing

    If intSize > 0 Then
        If Len(Content) > intSize Then

            If Len(Content) - intSize > 3 Then
            
                Content = Left(Content, intSize - 3) & "..."

            Else
            
                Content = Left(Content, intSize)

            End If
        End If
    End If
    
    ConvertHTML = Content
End Function

Public Sub WriteLog(ByVal str As String, Optional ByVal LogFileName As String = "", Optional ByVal bolNow As Boolean = True)
    On Error Resume Next
    Dim strLogFileName As String
    strLogFileName = "Log.txt"
    
    Dim iFileNum As Integer
        
    iFileNum = FreeFile()
    
    If LogFileName <> "" Then

        Open App.Path & "\log\" & LogFileName For Append As #iFileNum

    Else

        Open App.Path & "\log\" & strLogFileName For Append As #iFileNum

    End If

    If bolNow Then

        Print #iFileNum, str & "<-- " & Now()

    Else

        Print #iFileNum, str

    End If

    Close #iFileNum
    '.Visible = True

End Sub

'*************************************************************************
'**函 数 名： MySleep
'**输    入： DealyTime(Long) 需延时的时间
'**输    出： 无
'**功能描述： 经过改造的延时器无凝滞，突破原先TIMER控件65.5秒的限制
'*************************************************************************
Public Sub MySleep(DealyTime As Single)

    Dim TimerCount As Long

    TimerCount = Timer + DealyTime '增加N秒
    While TimerCount - Timer > 0

        DoEvents
        Sleep 10

        DoEvents
    Wend
End Sub


Public Sub SaveConfig(ByRef Frm As VB.Form)

        Dim ctl As Control
        Dim strConfig As String
        

100     For Each ctl In Frm.Controls
        
102         If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "RichTextBox" Or TypeName(ctl) = "ComboBox" Then
104             If ctl.Name <> "txtLog" Then
106                 strConfig = strConfig & ctl.Name & "|-|" & Replace(ctl.Text, vbCrLf, "|--|") & vbCrLf
                End If
108         ElseIf TypeName(ctl) = "CheckBox" Then
                
110                 strConfig = strConfig & ctl.Name & "|-|" & ctl.Value & vbCrLf

            End If
        
        Next
        
112     If strConfig <> "" Then

            Dim iFile As Integer

114         iFile = FreeFile()

116         Open App.Path & "\" & Frm.Name & ".Cfg" For Output As #iFile
118         Print #iFile, Left(strConfig, Len(strConfig) - 2)
120         Close #iFile
                    
        End If

End Sub

Public Sub ReadConfig(ByRef Frm As VB.Form)
        '<EhHeader>
        On Error GoTo ReadConfig_Err
        '</EhHeader>

        Dim iFile As Integer
    
100     iFile = FreeFile()

102     Open App.Path & "\" & Frm.Name & ".Cfg" For Input As #iFile
    
        'MsgBox App.Path
    
        Dim i As Integer
104     i = 1
    
    
108     Do While Not EOF(1)
            Dim strTmp As String
            Dim arr() As String
110         Line Input #iFile, strTmp
112         arr = Split(strTmp, "|-|", 2, vbBinaryCompare)

114         If UBound(arr) = 1 Then
        
116             CallByName Frm, arr(0), VbLet, Replace(arr(1), "|--|", vbCrLf)
        
118         Else
        
            End If
        
120         i = i + 1
        Loop

        Close #iFile

        '<EhFooter>
        Exit Sub

ReadConfig_Err:
        Close #iFile

        '</EhFooter>
End Sub
