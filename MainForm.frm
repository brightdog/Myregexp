VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "RegExp Test"
   ClientHeight    =   6090
   ClientLeft      =   1965
   ClientTop       =   2160
   ClientWidth     =   13530
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   13530
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDecodeURIUTF8 
      Caption         =   "UTF8 To Text"
      Height          =   285
      Left            =   4980
      TabIndex        =   33
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton cmdEncodeURIUTF8 
      Caption         =   "To UTF8"
      Height          =   285
      Left            =   4050
      TabIndex        =   32
      Top             =   1200
      Width           =   885
   End
   Begin VB.CommandButton cmdDecodeURIGB2312 
      Caption         =   "GB2312 To Text"
      Height          =   285
      Left            =   2490
      TabIndex        =   31
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdEncodeURIGB2312 
      Caption         =   "To GB2312"
      Height          =   285
      Left            =   1350
      TabIndex        =   30
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdURLList 
      Caption         =   "URL列表"
      Height          =   285
      Left            =   6330
      TabIndex        =   21
      Top             =   1230
      Width           =   885
   End
   Begin VB.CommandButton cmdURLRange 
      Caption         =   "连续URL"
      Height          =   285
      Left            =   7230
      TabIndex        =   20
      Top             =   1230
      Width           =   855
   End
   Begin VB.TextBox txtWebState 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9780
      TabIndex        =   19
      Top             =   1230
      Width           =   3045
   End
   Begin VB.ComboBox cboCharSet 
      Height          =   315
      Left            =   8700
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1230
      Width           =   1005
   End
   Begin VB.CommandButton cmdShowOpenURL 
      Caption         =   "URL"
      Height          =   285
      Left            =   8130
      TabIndex        =   17
      Top             =   1230
      Width           =   525
   End
   Begin VB.CommandButton cmdDecodeJson 
      Caption         =   "Decode Json"
      Height          =   285
      Left            =   90
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame fraTop 
      Caption         =   "Regular Expression"
      Height          =   1350
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   13305
      Begin VB.CommandButton cmdRestoreCrLf 
         Caption         =   "RestoreCrLf"
         Height          =   285
         Left            =   6840
         TabIndex        =   36
         Top             =   150
         Width           =   1605
      End
      Begin VB.CheckBox chkRestoreCRLF 
         Caption         =   "Restore CRLF"
         Height          =   315
         Left            =   12240
         TabIndex        =   35
         Top             =   900
         Width           =   975
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "Show History"
         Height          =   315
         Left            =   9030
         TabIndex        =   29
         Top             =   120
         Width           =   1305
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "添加"
         Height          =   285
         Left            =   5880
         TabIndex        =   28
         Top             =   150
         Width           =   705
      End
      Begin VB.ComboBox cboPattern 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   150
         Width           =   2745
      End
      Begin VB.CommandButton cmdReplaceSymbol 
         Caption         =   "替换保留字"
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   150
         Width           =   1125
      End
      Begin VB.ComboBox cboSplitTag 
         Height          =   300
         Left            =   9540
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   900
         Width           =   2565
      End
      Begin VB.CommandButton CopyPattern 
         Caption         =   "Copy Pattern"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   870
         Width           =   1485
      End
      Begin VB.CommandButton MatchPopUp 
         Caption         =   "Match-PopUp"
         Height          =   375
         Left            =   8250
         TabIndex        =   14
         Top             =   870
         Width           =   1215
      End
      Begin VB.CommandButton cmdbk 
         Caption         =   "Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4530
         TabIndex        =   13
         Top             =   870
         Width           =   510
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   870
         Width           =   975
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Colorize"
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   945
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "IgnoreCase"
         Height          =   240
         Left            =   1140
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkGlb 
         Caption         =   "Gloable"
         Height          =   240
         Left            =   105
         TabIndex        =   2
         Top             =   945
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.TextBox txtPattern 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Text            =   "(<(""[^""]*""|'[^']*'|[^'"">])*>)"
         Top             =   450
         Width           =   9645
      End
      Begin VB.CommandButton cmdMatch 
         Caption         =   "Match"
         Height          =   375
         Left            =   7200
         TabIndex        =   1
         Top             =   870
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pattern"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   495
         Width           =   510
      End
   End
   Begin TabDlg.SSTab tab 
      Height          =   5640
      Left            =   180
      TabIndex        =   4
      Top             =   1440
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   9948
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Text"
      TabPicture(0)   =   "MainForm.frx":00D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtSrc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDest"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkGroup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Matchs"
      TabPicture(1)   =   "MainForm.frx":00EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblWaiting"
      Tab(1).Control(1)=   "tvMatch"
      Tab(1).Control(2)=   "chkUseHtmlFilter"
      Tab(1).Control(3)=   "chkDecodeJson"
      Tab(1).Control(4)=   "cmdExpand"
      Tab(1).Control(5)=   "txtDetail"
      Tab(1).ControlCount=   6
      Begin VB.CheckBox chkGroup 
         Caption         =   "按照回车分组匹配（仅PopUp用）"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   60
         Width           =   3435
      End
      Begin RichTextLib.RichTextBox txtDetail 
         Height          =   1995
         Left            =   -73080
         TabIndex        =   34
         Top             =   1020
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3519
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"MainForm.frx":010A
      End
      Begin VB.CommandButton cmdExpand 
         Caption         =   "ExpandAll"
         Height          =   285
         Left            =   -70740
         TabIndex        =   24
         Top             =   0
         Width           =   1335
      End
      Begin VB.CheckBox chkDecodeJson 
         Caption         =   "Decode Json"
         Height          =   285
         Left            =   -72360
         TabIndex        =   23
         Top             =   0
         Width           =   1440
      End
      Begin VB.CheckBox chkUseHtmlFilter 
         Caption         =   "过滤html"
         Height          =   285
         Left            =   -73620
         TabIndex        =   12
         Top             =   0
         Width           =   1320
      End
      Begin MSComctlLib.TreeView tvMatch 
         Height          =   3795
         Left            =   -69600
         TabIndex        =   10
         Top             =   810
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   6694
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtDest 
         Height          =   5865
         Left            =   4905
         TabIndex        =   6
         Top             =   495
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10345
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"MainForm.frx":01A7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtSrc 
         Height          =   5820
         Left            =   225
         TabIndex        =   5
         Top             =   480
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   10266
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"MainForm.frx":0234
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblWaiting 
         Alignment       =   2  'Center
         Caption         =   $"MainForm.frx":02D1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74460
         TabIndex        =   25
         Top             =   780
         Width           =   4245
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txtReplaceBack As String

Dim Mc  As Object
Dim col As New Collection
Dim strURL As String
Dim strPost As String
Dim strReferer As String
Public strURLList As String
Public strURLBegin As String
Dim strLastPattern As String
Private m_regexp        As Object

Private Sub cmbType_Click()
    Set m_regexp = New VBScript_RegExp_55.RegExp

End Sub

Private Sub cmdAppend_Click()
    Dim iStart As Integer
    Dim iEnd As Integer
    
    iStart = Me.txtPattern.SelStart
    iEnd = Me.txtPattern.SelStart + Me.txtPattern.SelLength
    
    Me.txtPattern.Text = Left(Me.txtPattern.Text, iStart) & Me.cboPattern.List(Me.cboPattern.ListIndex) & Right(Me.txtPattern.Text, Len(Me.txtPattern.Text) - iEnd)
    
End Sub

Private Sub cmdbk_Click()
    txtSrc.Text = txtReplaceBack
    txtReplaceBack = ""
    cmdbk.Enabled = False
End Sub

Private Sub cmdDecodeJson_Click()
    Dim objDecode As New clsEncodeURI
    Me.txtDest.Text = objDecode.Unicode_Decode(Me.txtSrc.Text)
    Set objDecode = Nothing
End Sub

Private Sub cmdDecodeURIGB2312_Click()
        '<EhHeader>
        On Error GoTo cmdDecodeURIGB2312_Click_Err
        '</EhHeader>
'100     If Len(Me.txtSrc.Text) > 1000 Then
'102         MsgBox "Too many Chars!!"
'            Exit Sub
'        End If
      
        Dim i As Integer
    
        Dim strTmp As String
        Dim strResult As String
104     strTmp = Me.txtSrc.Text
    
    
106     For i = 1 To Len(strTmp)
    
    
108         If Mid(strTmp, i, 1) = "%" Then '是GB2312编码
                
                Dim strNextChar As String
                strNextChar = Mid(strTmp, i + 1, 1)
                Select Case Val(strNextChar)
                
                
                Case Is > 0 '英文符号
                
                    strResult = strResult & Chr(Val("&H " & Mid(strTmp, i + 1, 2)))
                    i = i + 2
                Case Else  '中文
                    
                    strResult = strResult & Chr(Val("&H " & Mid(strTmp, i + 1, 2) & Mid(strTmp, i + 4, 2)))
                    i = i + 5
                End Select
        
            Else
            
                strResult = strResult & Mid(strTmp, i, 1)
                
            End If
    
    
    
        Next
        Me.txtDest.Text = strResult
        '<EhFooter>
        Exit Sub

cmdDecodeURIGB2312_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub cmdDecodeURIUTF8_Click()
        '<EhHeader>
        On Error GoTo cmdDecodeURIUTF8_Click_Err

        '</EhHeader>
'100     If Len(Me.txtSrc.Text) > 200 Then
'
'102         MsgBox "Too many Chars!!"
'            Exit Sub
'
'        End If

        On Error Resume Next
        Dim obj As clsEncodeURI
104     Set obj = New clsEncodeURI
    
106     Me.txtDest.Text = obj.UTF8ToGB2312(Me.txtSrc.Text)
108     Set obj = Nothing
        '<EhFooter>
        Exit Sub

cmdDecodeURIUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & _
           "at line " & Erl

        '</EhFooter>
End Sub

Private Sub cmdEncodeURIGB2312_Click()
        '<EhHeader>
        On Error GoTo cmdEncodeURIGB2312_Click_Err
        '</EhHeader>
100     If Len(Me.txtSrc.Text) > 200 Then
102         MsgBox "Too many Chars!!"
            Exit Sub
        End If
        On Error Resume Next
        Dim obj As clsEncodeURI
104     Set obj = New clsEncodeURI
    
106     Me.txtDest.Text = obj.ChineseToGB2312(Me.txtSrc.Text)
108     Set obj = Nothing
        '<EhFooter>
        Exit Sub

cmdEncodeURIGB2312_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "at line " & Erl
        '</EhFooter>
End Sub

Private Sub cmdEncodeURIUTF8_Click()
        '<EhHeader>
        On Error GoTo cmdEncodeURIUTF8_Click_Err
        '</EhHeader>
100     If Len(Me.txtSrc.Text) > 200 Then
102         MsgBox "Too many Chars!!"
            Exit Sub
        End If
        On Error Resume Next
        Dim obj As clsEncodeURI
104     Set obj = New clsEncodeURI
    
106     Me.txtDest.Text = obj.ChineseToUTF8(Me.txtSrc.Text)
108     Set obj = Nothing
        '<EhFooter>
        Exit Sub

cmdEncodeURIUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "at line " & Erl
        '</EhFooter>
End Sub

Private Sub cmdExpand_Click()
    Dim bol As Boolean
    bol = False
    
    If Me.cmdExpand.Caption = "ExpandAll" Then

        bol = True
        Me.cmdExpand.Caption = "CollapseAll"

    Else

        bol = False
        Me.cmdExpand.Caption = "ExpandAll"

    End If

    Dim i As Long
    
    Me.tvMatch.Visible = False
    
    For i = 1 To Me.tvMatch.Nodes.Count
            
        Me.tvMatch.Nodes.Item(i).Expanded = bol
        DoEvents
    Next
    
    Me.tvMatch.Visible = True
    If Me.tvMatch.SelectedItem Is Nothing And bol And Me.tvMatch.Nodes.Count > 0 Then
        
        Me.tvMatch.Nodes.Item(1).Selected = True
        
    End If
    Debug.Print Me.tvMatch.SelectedItem.Visible
    'Me.tvMatch.Nodes.Item(Me.tvMatch.SelectedItem.index).Selected = True
    
    Me.tvMatch.SetFocus
    
End Sub


Private Sub cmdHistory_Click()
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If fso.FileExists(App.Path & "\History.pattern") Then
        
        
        Dim strHistory As String
        strHistory = fso.GetFile(App.Path & "\History.pattern").OpenAsTextStream.ReadAll
        
        Dim frmHis As frmHistory
        Set frmHis = New frmHistory
        
        frmHis.strHistory = strHistory
        frmHis.Show
        
    End If
    
    Set fso = Nothing
End Sub

Private Sub SavePatternToHistory()
    
    If Me.txtPattern.Text <> "" Then

        If strLastPattern <> Me.txtPattern.Text Then

            strLastPattern = Me.txtPattern.Text
            Open App.Path & "\History.pattern" For Append As #1
            Print #1, Me.txtPattern.Text
            Close #1

        End If

    End If

End Sub

Private Sub cmdMatch_Click()
    Call SaveConfig(Me)
    Call SavePatternToHistory
    strLastPattern = Me.txtPattern.Text
    Me.cmdExpand.Caption = "ExpandAll"

    Me.txtDest.SelStart = 0
    Me.txtDest.SelLength = Len(Me.txtDest.Text)
    Me.txtDest.SelColor = vbBlack
    Me.txtDest.SelBold = False
    
    Dim m       As Variant
    Dim pos     As Long
    Dim i       As Integer
    Dim Value   As String
    Dim s       As String
    Dim a()     As String
    Dim v
    Dim nodX    As Node
    Dim nodParent   As Node

    Me.txtDest.SelColor = vbBlack
 
    tvMatch.Nodes.Clear
    txtDetail.Text = ""
    
    If Len(Me.txtSrc.Text) > 100000 And Me.chkShow.Value = 1 Then
        If MsgBox("too much chars, do you want to disable colorizing?", vbQuestion + vbYesNo) = vbYes Then
            Me.chkShow.Value = 0
        End If
    End If
    
    m_regexp.IgnoreCase = IIf(Me.chkCase.Value = 1, True, False)
    m_regexp.Global = IIf(Me.chkGlb.Value = 1, True, False)
 
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, vbCr, "", 1, -1, vbBinaryCompare)
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, vbLf, "", 1, -1, vbBinaryCompare)
    
 
    m_regexp.Pattern = Me.txtPattern.Text

    s = modCRLF.convertCRLF(Me.txtSrc.Text)
    
    On Error GoTo regexp_error
    Set Mc = m_regexp.Execute(s)
    'On Error GoTo 0
    
    Set col = New Collection
    
    
    Me.tvMatch.Visible = False
    For Each m In Mc

        col.Add m.Value
        
        
        AddMatchToTree Nothing, m
        
        
        DoEvents
         
    Next m
    Me.tvMatch.Visible = True
    Me.Caption = "RegExp Test   -   " & col.Count
    Me.txtDest.Text = Me.txtSrc.Text

    If col.Count = 0 Then
        Exit Sub

    End If
    
    s = Me.txtDest.Text
    
    If Me.chkShow.Value = 1 Then
        Me.txtDest.Visible = False

        i = 1

        Do While True

            If i > col.Count Then
                Exit Do
            End If

            Value = col.Item(i)

            pos = InStr(pos + 1, s, restoreCRLF(Value), vbBinaryCompare)

            If pos > 0 Then
                Me.txtDest.SelStart = pos - 1
                Me.txtDest.SelLength = Len(Value)
                Me.txtDest.SelColor = vbRed
                Me.txtDest.SelBold = True
                'Me.txtDest.SelText = Value
            Else
                Exit Do
            End If

            i = i + 1

        Loop

        Me.txtDest.Visible = True
    End If
    
    
    
    Exit Sub
    
regexp_error:
    Me.txtDest.Visible = True
    MsgBox "Match error:" & Err.Description, vbCritical + vbOKOnly
    
End Sub
 
Private Sub AddMatchToTree(ByRef nodParent As Node, ByRef m As Variant)
    Dim nodX    As Node
    Dim sm      As Variant

    On Error Resume Next
    
    If nodParent Is Nothing Then
        Set nodX = tvMatch.Nodes.Add(, , , m)
    Else
        Set nodX = tvMatch.Nodes.Add(nodParent, tvwChild, , m)
    End If
    Dim i As Integer
    i = 0
    For Each sm In m.SubMatches
        Call tvMatch.Nodes.Add(nodX, tvwChild, , "(" & i & ")-" & sm)
        Debug.Print sm
        i = i + 1
    Next sm
    DoEvents
End Sub

Private Sub cmdReplace_Click()
    txtReplaceBack = txtSrc.Text
    Dim regTmp As VBScript_RegExp_55.RegExp
    Set regTmp = New VBScript_RegExp_55.RegExp
    regTmp.Global = True
    regTmp.MultiLine = True
    regTmp.IgnoreCase = True
    regTmp.Pattern = Me.txtPattern.Text
    txtSrc.Text = restoreCRLF(regTmp.Replace(convertCRLF(txtSrc.Text), ""))
    Set regTmp = Nothing
    cmdbk.Enabled = True
End Sub

Private Function isURL(ByVal strURL As String) As Boolean
    isURL = False
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    Reg.Global = True
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    Reg.Pattern = "^(http|https)\://([a-zA-Z0-9\.\-]+(\:[a-zA-Z0-9\.&%\$\-]+)*@)*((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])|localhost|([a-zA-Z0-9\-]+\.)*[a-zA-Z0-9\-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{1,10}))(\:[0-9]+)*(/($|[a-zA-Z0-9;\.\,\?\'\\\+&%\$#\=~_\-*\|]+))*$"
    isURL = Reg.Test(strURL)
End Function

Private Sub cmdReplaceSymbol_Click()
    
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, """""", """")
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, "\?", "?")
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, "?", "\?")
    Me.txtPattern.Text = Replace(Me.txtPattern.Text, ".*\?", ".*?")
    
End Sub

Private Sub cmdRestoreCrLf_Click()
    Me.txtSrc.Text = restoreCRLF(Me.txtSrc.Text)
End Sub

Private Sub cmdShowOpenURL_Click()
    
reOpen:
    Dim MsgForm As frmURL
    Set MsgForm = New frmURL
    Dim bolNeedLoadUrl As Boolean
    'Call frmURL.ILoade(Me, strURL)
    MsgForm.lastURL = strURL
    MsgForm.lastPost = strPost
    MsgForm.txtReferer.Text = strReferer
    MsgForm.Show vbModal, Me
    bolNeedLoadUrl = MsgForm.btnOK

    If bolNeedLoadUrl Then

        strURL = MsgForm.URL
        strPost = MsgForm.POST
        strReferer = MsgForm.txtReferer.Text

        If strURL = "" Then

            MsgBox ("请输入URL")
            GoTo reOpen

        End If

    End If
    
    Unload MsgForm
    Set MsgForm = Nothing

    If bolNeedLoadUrl Then

        If LCase(Left(strURL, 7)) <> "http://" And LCase(Left(strURL, 8)) <> "https://" Then

            strURL = "http://" & strURL

        End If

        'If isURL(strURL) Then
            Dim iWeb As Object

            If iWeb Is Nothing Then

                'Dim iWeb As clsXMLHTTPGetHtml
                If strReferer <> "" Then
                    Set iWeb = New clsWinHTTPGetHtml
                Else
                
                    Set iWeb = New clsXMLHTTPGetHtml
                End If
            End If

            iWeb.URL = strURL
            iWeb.PostData = strPost
            iWeb.Referer = strReferer
            iWeb.CharSet = Me.cboCharSet.List(Me.cboCharSet.ListIndex)
            Dim strResult As String
            strResult = restoreCRLF(iWeb.StartGetHtml)
        
            Me.txtSrc.Text = strResult
            Set iWeb = Nothing

'        Else
'
'            If strURL <> "" Then
'
'                MsgBox "输入的URL不合法，请重新输入！"
'                GoTo reOpen
'
'            End If
'
'        End If

    End If

End Sub

Private Sub cmdURLList_Click()
    Dim MsgForm As frmURLList
    Set MsgForm = New frmURLList
    'Load frmURLList
    MsgForm.LastURLList = strURLList
    MsgForm.LastURLBegin = strURLBegin
    MsgForm.Show
End Sub

Private Sub cmdURLRange_Click()
    Load frmURLRange
    frmURLRange.Show
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Set m_regexp = New VBScript_RegExp_55.RegExp
102     Me.tab.tab = 0
104     txtSrc.ZOrder 0
106     txtDest.ZOrder 0
    
108     Me.cboCharSet.AddItem "GB2312"
110     Me.cboCharSet.AddItem "UTF-8"
112     Me.cboCharSet.AddItem "GBK"
114     Me.cboCharSet.ListIndex = 1
    
116     Me.cboSplitTag.AddItem "vbTab"
118     Me.cboSplitTag.AddItem "<br style=""clear: both;"" />"
120     Me.cboSplitTag.AddItem "<hr />"
122     Me.cboSplitTag.AddItem "|--|"
124     Me.cboSplitTag.AddItem "|"
    
        Dim fso As Scripting.FileSystemObject
126     Set fso = New Scripting.FileSystemObject
    
128     If fso.FileExists(App.Path & "\src.text") Then
130         Me.txtSrc.Text = fso.GetFile(App.Path & "\src.text").OpenAsTextStream.ReadAll
132         Me.txtSrc.Text = Left(Me.txtSrc.Text, Len(Me.txtSrc.Text) - 2)
        End If
    
134     If fso.FileExists(App.Path & "\pattern.text") Then
136         Me.txtPattern.Text = fso.GetFile(App.Path & "\pattern.text").OpenAsTextStream.ReadLine
        End If
    
138     If fso.FileExists(App.Path & "\url.text") Then
140         strURL = fso.GetFile(App.Path & "\url.text").OpenAsTextStream.ReadLine
        End If

142     If fso.FileExists(App.Path & "\post.text") Then
144         strPost = fso.GetFile(App.Path & "\post.text").OpenAsTextStream.ReadLine
        End If

146     If fso.FileExists(App.Path & "\referer.text") Then
148         strReferer = fso.GetFile(App.Path & "\referer.text").OpenAsTextStream.ReadLine
        End If

150     If fso.FileExists(App.Path & "\urllist.text") Then
152         strURLList = fso.GetFile(App.Path & "\urllist.text").OpenAsTextStream.ReadAll
        End If

154     If fso.FileExists(App.Path & "\urllist_begin.text") Then
156         strURLBegin = fso.GetFile(App.Path & "\urllist_begin.text").OpenAsTextStream.ReadAll
        End If
    
158     If fso.FileExists(App.Path & "\model.pattern") Then

            Dim arrPattern() As String
160         arrPattern = Split(fso.GetFile(App.Path & "\model.pattern").OpenAsTextStream.ReadAll & vbCrLf, vbCrLf, -1, vbBinaryCompare)

162         If UBound(arrPattern) > 0 Then
                Dim i As Integer
            
164             For i = 0 To UBound(arrPattern)
            
166                 If arrPattern(i) <> "" Then
                
168                     Me.cboPattern.AddItem arrPattern(i)
                
                    End If
            
                Next
        
            End If
        End If

        Call ReadConfig(Me)
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteLog Err.Description & vbCrLf & "in RegExpTest.MainForm.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Open App.Path & "\src.text" For Output As #1
    Print #1, Me.txtSrc.Text
    Close #1
    
    Open App.Path & "\pattern.text" For Output As #1
    Print #1, Me.txtPattern.Text
    Close #1
    
    If strURL <> "" Then

        Open App.Path & "\url.text" For Output As #1
        Print #1, strURL
        Close #1

    End If

    Open App.Path & "\post.text" For Output As #1
    Print #1, strPost
    Close #1
    
    Open App.Path & "\referer.text" For Output As #1
    Print #1, strReferer
    Close #1

    If strURLList <> "" Then

        Open App.Path & "\urllist.text" For Output As #1
        Print #1, strURLList
        Close #1

    End If
    
    If strURLBegin <> "" Then

        Open App.Path & "\urllist_begin.text" For Output As #1
        Print #1, strURLBegin
        Close #1

    End If
    
    Call SaveConfig(Me)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.fraTop.Move 100, 50, width - 300, fraTop.Height
    Me.tab.Move 100, 1800, width - 300, Height - 2200
    
    Me.cmdDecodeJson.Top = Me.fraTop.Height + 80
    Me.cmdURLList.Top = Me.cmdDecodeJson.Top
    Me.cmdURLRange.Top = Me.cmdDecodeJson.Top
    Me.cmdShowOpenURL.Top = Me.cmdDecodeJson.Top
    Me.cboCharSet.Top = Me.cmdDecodeJson.Top
    Me.txtWebState.Top = Me.cmdDecodeJson.Top
    
    Me.cmdEncodeURIGB2312.Top = Me.cmdDecodeJson.Top
    Me.cmdEncodeURIUTF8.Top = Me.cmdDecodeJson.Top
    Me.cmdDecodeURIGB2312.Top = Me.cmdDecodeJson.Top
    Me.cmdDecodeURIUTF8.Top = Me.cmdDecodeJson.Top
    
    Me.txtSrc.Move 50, 350, (Me.tab.width - 200) / 2, Me.tab.Height - 400
    Me.txtDest.Move txtSrc.Left + txtSrc.width + 50, txtSrc.Top, txtSrc.width, txtSrc.Height
    Me.lblWaiting.Move txtSrc.Left, txtSrc.Top, txtSrc.width, txtSrc.Height
    Me.tvMatch.Move txtSrc.Left, txtSrc.Top, txtSrc.width, txtSrc.Height
    Me.txtDetail.Move txtDest.Left, txtDest.Top, txtDest.width, txtDest.Height
    Me.txtPattern.width = Me.width - 1260

End Sub

Private Sub CopyPattern_Click()
    Call Clipboard.Clear
    Call Clipboard.SetText("""" & Replace(Me.txtPattern.Text, """", """""") & """")
End Sub

Private Sub MatchPopUp_Click()
    Call SaveConfig(Me)
    Call SavePatternToHistory
    
    Dim myReg As New RegExp

    Dim Mc As MatchCollection

    myReg.IgnoreCase = True
    myReg.Global = True

    myReg.MultiLine = False
    Dim m As Variant
    Dim sm As Variant

    Dim i As Double
    Dim j As Double
    Dim strSplitTag As String
    txtPattern.Text = Replace(txtPattern.Text, vbCrLf, "", 1, -1, vbTextCompare)

    myReg.Pattern = txtPattern.Text
    Dim objDecode As clsEncodeURI
    frmPop.txtResault.Text = ""
    Dim SB As clsStringBuilder
    Dim tmpi As Double
    Dim tmpStr As String
    Dim strResult As String

    If chkGroup.Value = 1 Then
        
        Dim arrLine() As String
        
        arrLine = Split(Me.txtSrc.Text, vbCrLf, -1, vbBinaryCompare)
        
        If UBound(arrLine) > -1 Then
        
            Dim l As Integer
            
            For l = 0 To UBound(arrLine)
                Set Mc = myReg.Execute(convertCRLF(arrLine(l)))

                i = Mc.Count

                j = 0
        
                strSplitTag = Me.cboSplitTag.List(Me.cboSplitTag.ListIndex)

                If strSplitTag = "vbTab" Then
                    strSplitTag = vbTab
                End If

                Set SB = New clsStringBuilder
                
                Set objDecode = New clsEncodeURI
        
                For Each m In Mc

                    tmpi = 1
            
                    For Each sm In m.SubMatches

                        tmpStr = CStr(sm)

                        If Me.chkUseHtmlFilter.Value = 1 Then
                            tmpStr = ConvertHTML.ConvertHTML(tmpStr)
                        End If
                
                        If Me.chkDecodeJson.Value = 1 Then
                            tmpStr = objDecode.Unicode_Decode(tmpStr)
                        End If
                
                        If tmpi = m.SubMatches.Count Then
                            SB.Append tmpStr
                        Else
                            SB.Append tmpStr & strSplitTag
                        End If

                        tmpi = tmpi + 1
                        j = j + 1
                        Me.Caption = "Main结果： " & j & "/" & i & " 行"

                        DoEvents
                    Next

                    SB.Append vbCrLf

                Next

                Set objDecode = Nothing
                frmPop.Caption = "Pop结果： " & Me.Caption & " 行"
        
                strResult = Replace(SB.ToString, vbTab & vbCrLf, vbCrLf, 1, -1, vbBinaryCompare)
                frmPop.txtResault.Text = frmPop.txtResault.Text & vbCrLf & "<LINE: " & l + 1 & " START>" & vbCrLf

                If Me.chkRestoreCRLF.Value = 1 Then
        
                    frmPop.txtResault.Text = frmPop.txtResault.Text & restoreCRLF(strResult)
                Else
        
                    frmPop.txtResault.Text = frmPop.txtResault.Text & strResult
                End If

                frmPop.txtResault.Text = frmPop.txtResault.Text & vbCrLf & "<LINE: " & l + 1 & " END>" & vbCrLf
                frmPop.Show
            
            Next
        
        End If
        
    Else

        If myReg.Test(convertCRLF(txtSrc.Text)) Then

            Set Mc = myReg.Execute(convertCRLF(txtSrc.Text))

            i = Mc.Count

            j = 0
        
            strSplitTag = Me.cboSplitTag.List(Me.cboSplitTag.ListIndex)

            If strSplitTag = "vbTab" Then
                strSplitTag = vbTab
            End If

            Set SB = New clsStringBuilder

            Set objDecode = New clsEncodeURI
        
            For Each m In Mc
                
                tmpi = 1
            
                For Each sm In m.SubMatches
                    
                    tmpStr = CStr(sm)

                    If Me.chkUseHtmlFilter.Value = 1 Then
                        tmpStr = ConvertHTML.ConvertHTML(tmpStr)
                    End If
                
                    If Me.chkDecodeJson.Value = 1 Then
                        tmpStr = objDecode.Unicode_Decode(tmpStr)
                    End If
                
                    If tmpi = m.SubMatches.Count Then
                        SB.Append tmpStr
                    Else
                        SB.Append tmpStr & strSplitTag
                    End If

                    tmpi = tmpi + 1
                    j = j + 1
                    Me.Caption = "已完成 " & j & "/" & i & " 个"

                    DoEvents
                Next

                SB.Append vbCrLf

            Next

            Set objDecode = Nothing
            frmPop.Caption = "匹配结果： " & Me.Caption
        
            strResult = Replace(SB.ToString, vbTab & vbCrLf, vbCrLf, 1, -1, vbBinaryCompare)
        
            If Me.chkRestoreCRLF.Value = 1 Then
        
                frmPop.txtResault.Text = restoreCRLF(strResult)
            Else
        
                frmPop.txtResault.Text = strResult
            End If
        
            frmPop.Show
    
        Else

            MsgBox "表达式错误，请检查！"

        End If
    End If


    Set SB = Nothing

End Sub

Private Sub tab_Click(PreviousTab As Integer)

    If Me.tab.tab = 0 Then
        txtSrc.ZOrder 0
        txtDest.ZOrder 0
    Else
        lblWaiting.ZOrder 0
        tvMatch.ZOrder 0
        txtDetail.ZOrder 0
    End If
    
End Sub


Private Sub tvMatch_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim strTmp As String
    
    strTmp = Node.Text
    If Me.chkUseHtmlFilter.Value = 1 Then
        strTmp = ConvertHTML.ConvertHTML(Node.Text)
    End If
    If Me.chkDecodeJson.Value = 1 Then
        Dim objDecode As New clsEncodeURI
        strTmp = objDecode.Unicode_Decode(strTmp)
        Set objDecode = Nothing
    End If
    
    
    txtDetail.Text = restoreCRLF(strTmp)


End Sub

Private Sub txtDest_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        Me.txtDest.SelStart = 0
        Me.txtDest.SelLength = Len(Me.txtDest.Text)
    End If
End Sub

Private Sub txtSrc_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        Me.txtSrc.SelStart = 0
        Me.txtSrc.SelLength = Len(Me.txtSrc.Text)
    End If
End Sub


