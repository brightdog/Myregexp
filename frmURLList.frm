VERSION 5.00
Begin VB.Form frmURLList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "URLList"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmURLList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7755
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtURL 
      Height          =   6075
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   300
      Width           =   7635
   End
   Begin VB.ComboBox cboMethod 
      Height          =   300
      ItemData        =   "frmURLList.frx":000C
      Left            =   660
      List            =   "frmURLList.frx":0016
      TabIndex        =   13
      Top             =   0
      Width           =   1155
   End
   Begin VB.CheckBox chkDirectGetResponseText 
      Caption         =   "直接取ResponseText"
      Height          =   195
      Left            =   4500
      TabIndex        =   12
      Top             =   60
      Width           =   3195
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   6960
      Width           =   5295
   End
   Begin VB.CheckBox chkWriteFile 
      Caption         =   "写文件"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7020
      Width           =   1695
   End
   Begin VB.CheckBox chkUrlEncode 
      Caption         =   "需要编码"
      Height          =   195
      Left            =   3180
      TabIndex        =   9
      Top             =   60
      Width           =   1875
   End
   Begin VB.ComboBox cboCharSet 
      Height          =   300
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   0
      Width           =   1005
   End
   Begin VB.TextBox txtUrlBegin 
      Height          =   375
      Left            =   1260
      TabIndex        =   6
      Top             =   6420
      Width           =   6435
   End
   Begin VB.CheckBox chkDecodeJson 
      Caption         =   "Decode Json"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   7860
      Width           =   1440
   End
   Begin VB.CheckBox chkAppend 
      Caption         =   "附加到主窗体文本之后"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   2355
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3870
      TabIndex        =   3
      Top             =   7710
      Width           =   1005
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4890
      TabIndex        =   2
      Top             =   7710
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5880
      TabIndex        =   1
      Top             =   7710
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "前缀/Action"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   6540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Method"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   435
   End
End
Attribute VB_Name = "frmURLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolPause As Boolean
Private bolStop As Boolean

Public LastURLList As String
Public LastURLBegin As String

Private Sub cmdOK_Click()
100     Call SaveConfig(Me)
102     bolPause = False
104     bolStop = False
    
        Dim bolHasPaused As Boolean
106     bolHasPaused = False
    
108     Me.cmdOK.Enabled = False
110     Me.txtURL.Enabled = False
    
112     Me.cmdPause.Enabled = True
114     Me.cmdStop.Enabled = True
    
        Dim strURL As String
    
        Dim i      As Long
        Dim iWeb   As clsXMLHTTPGetHtml
116     Set iWeb = New clsXMLHTTPGetHtml
    
        'iWeb.DirectGetResponseText = True
    
        Dim SB As clsStringBuilder
118     Set SB = New clsStringBuilder
    
        Dim arrURL() As String
        Dim arrPost() As String
    
120     arrURL = Split(Me.txtURL.Text & vbCrLf, vbCrLf, -1, vbBinaryCompare)
    
        Dim ListCnt As Long
    
122     ListCnt = UBound(arrURL)
    
124     For i = 0 To ListCnt
        
126         If bolPause Then

128             If Me.chkAppend = 0 And Not bolHasPaused Then

130                 MainForm.txtSrc.Text = ""

                End If

132             bolHasPaused = True

134             If Me.chkWriteFile.Value = 1 Then
                
                    '                Dim intFile As Integer
                    '                intFile = VBA.FreeFile()
                    '
                    '                Open App.Path & "\" & Me.txtFileName.Text For Append As #intFile
                    '                Print #intFile, SB.ToString
                    '                Close #intFile
                
                Else
136                 MainForm.txtSrc.Text = MainForm.txtSrc.Text & SB.ToString
                End If

138             MainForm.strURLList = Me.txtURL.Text
140             SB.Value = ""
            
                Do

142                 MySleep 0.5

144                 If bolStop Then
        
                        Exit For

                    End If

146             Loop While bolPause

            End If
        
148         If bolStop Then
        
                Exit For

            End If
        
150         If arrURL(i) <> "" Then

                Dim objDecode As New clsEncodeURI
152             Me.Caption = i & ":" & ListCnt & " : " & arrURL(i)
            
154             If Me.txtUrlBegin.Text = "" Then
            
156                 If LCase(Left(arrURL(i), 7)) <> "http://" Then

158                     strURL = "http://" & arrURL(i)
                    Else
160                     strURL = arrURL(i)
                    End If
iWeb.URL = strURL
                Else
                
162                 If Me.cboMethod.Text = "GET" Then
                
164                     strURL = Me.txtUrlBegin.Text & arrURL(i)
    
166                     If LCase(Left(strURL, 7)) <> "http://" Then
    
168                         strURL = "http://" & strURL
    
                        End If
                        
170                     If Me.chkUrlEncode.Value = 1 Then
    
172                         Select Case Me.cboCharSet.Text
    
                                Case "GB2312"
174                                 strURL = Me.txtUrlBegin.Text & objDecode.ChineseToGB2312(arrURL(i))
    
176                             Case "UTF-8"
178                                 strURL = Me.txtUrlBegin.Text & objDecode.ChineseToUTF8(arrURL(i))
    
                            End Select
    
                        End If

180                     iWeb.URL = strURL
                    Else
                    
182                     arrPost = arrURL

184                     strURL = Me.txtUrlBegin.Text
    
186                     If LCase(Left(strURL, 7)) <> "http://" Then
    
188                         strURL = "http://" & strURL
    
                        End If

190                     iWeb.URL = strURL
192                     iWeb.PostData = arrURL(i)
                    End If
                End If
            

194             WriteLog "##:" & arrURL(i) & "|**|" & strURL
196             iWeb.CharSet = Me.cboCharSet.Text
                Dim strResult As String

198             If Me.cboMethod.ListIndex = 1 Then
200                 If UBound(arrPost) >= 0 Then
                        Dim j As Integer
                        Dim SBtmp As clsStringBuilder
202                     Set SBtmp = New clsStringBuilder
204                     SBtmp.Value = ""

206                     For j = 0 To UBound(arrPost)
208                         Me.Caption = j & ":" & ListCnt & " : " & arrPost(j)

210                         If arrPost(j) <> "" Then
212                             iWeb.URL = Me.txtUrlBegin.Text
214                             iWeb.PostData = arrPost(j)
                        
216                             SBtmp.Append restoreCRLF(iWeb.StartGetHtml)
                            End If

                        Next
                
218                     strResult = SBtmp.ToString
220                     Set SBtmp = Nothing
                        Exit For
                    End If

                Else
222                 strResult = restoreCRLF(iWeb.StartGetHtml)
                End If

224             If Me.chkDecodeJson.Value = 1 Then
                
226                 strResult = objDecode.Unicode_Decode(strResult)
                
                End If

228             WriteLog "@@:" & strResult

230             If Me.chkWriteFile.Value = 1 Then
                
                    Dim intFile As Integer
232                 intFile = VBA.FreeFile()
                    Dim strPath As String

234                 If InStr(1, Me.txtFileName.Text, ":", vbBinaryCompare) > 0 Then
236                     strPath = Me.txtFileName.Text
                    Else
238                     strPath = App.Path & "\" & Me.txtFileName.Text
                    End If

240                 Open strPath For Append As #intFile
242                 Print #intFile, strResult
244                 Close #intFile
                Else
246                 SB.Append strResult & vbCrLf
                End If

248             Set objDecode = Nothing
            End If

250         DoEvents

        Next

252     If i = 0 And Me.chkAppend = 0 And Not bolHasPaused Then

254         MainForm.txtSrc.Text = ""

        End If

256     MainForm.txtSrc.Text = SB.ToString
258     MainForm.strURLList = Me.txtURL.Text
    
260     Set iWeb = Nothing
262     Unload Me
End Sub

Private Sub cmdPause_Click()

    If Not bolPause Then

        Me.cmdPause.Caption = "Continue"
        bolPause = True

    Else

        Me.cmdPause.Caption = "Pause"
        bolPause = False

    End If

End Sub

Private Sub cmdStop_Click()
    bolStop = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then

        Unload Me

    End If

End Sub

Private Sub Form_Load()
    Me.txtURL.Text = LastURLList
    Me.txtURL.SelStart = 0
    Me.txtURL.SelLength = Len(Me.txtURL.Text)
    Me.cboCharSet.AddItem "GB2312"
    Me.cboCharSet.AddItem "UTF-8"
    Me.cboCharSet.ListIndex = 1
    Call ReadConfig(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveConfig(Me)
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    If Me.txtFileName.Text <> "" Then
        Me.chkWriteFile.Value = 1
    End If
End Sub

Private Sub txtURL_Change()

    If Trim(Me.txtURL.Text) <> "" Then

        Me.cmdOK.Enabled = True

    End If

End Sub

Private Sub txtURL_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then

        Unload Me

    End If

End Sub
