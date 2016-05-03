VERSION 5.00
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History - Only Last 100 History Records, for more history please open History.pattern By Notepad"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   15165
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All"
      Height          =   495
      Left            =   6540
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ListBox lstHistory 
      Height          =   3840
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15045
   End
   Begin VB.Label lblCnt 
      Height          =   435
      Left            =   4380
      TabIndex        =   2
      Top             =   4020
      Width           =   1935
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strHistory As String

Private Sub cmdShowAll_Click()
    Call LoadHistoryPattern(-1)
End Sub

Private Sub LoadHistoryPattern(Optional ByRef MaxCnt As Integer = 100)

    Dim i As Integer
    Dim arrHistory() As String
    arrHistory = Split(strHistory & vbCrLf, vbCrLf, -1, vbBinaryCompare)

    Dim iCnt As Long
    Dim iMaxCnt As Long

    If MaxCnt = -1 Then
        iMaxCnt = 100000
    Else
        iMaxCnt = MaxCnt
    End If
    
    iCnt = 1
    
    For i = UBound(arrHistory) To 0 Step -1

        If iCnt <= iMaxCnt Then
            If arrHistory(i) <> "" Then
    
                Me.lstHistory.AddItem arrHistory(i)

                If MaxCnt >= 0 Then
                    Me.lblCnt = iCnt & " / " & iMaxCnt
                Else
                    Me.lblCnt = iCnt
                End If

                iCnt = iCnt + 1
                
            End If

        Else
        
            Exit For
            
        End If

        DoEvents
    Next

End Sub

Private Sub Form_Load()
    Call LoadHistoryPattern
End Sub

Private Sub lstHistory_DblClick()
    MainForm.txtPattern.Text = Me.lstHistory.List(Me.lstHistory.ListIndex)
    Unload Me
End Sub

Private Sub lstHistory_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then

        Unload Me

    End If

End Sub

