VERSION 5.00
Begin VB.Form frmURLRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "URLRange"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmURLRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   6405
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   435
      Left            =   4530
      TabIndex        =   7
      Top             =   660
      Width           =   1755
   End
   Begin VB.TextBox txtEnd 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   750
      Width           =   495
   End
   Begin VB.TextBox txtStart 
      Height          =   315
      Left            =   810
      TabIndex        =   2
      Top             =   750
      Width           =   495
   End
   Begin VB.TextBox txtURL 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      Top             =   60
      Width           =   5475
   End
   Begin VB.Label Label4 
      Caption         =   "URL中需要连续变换的部分用(*)替换"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   810
      TabIndex        =   6
      Top             =   450
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "-->"
      Height          =   255
      Left            =   1380
      TabIndex        =   4
      Top             =   780
      Width           =   285
   End
   Begin VB.Label Label2 
      Caption         =   "范围"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   780
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmURLRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Call SaveConfig(Me)
    If InStr(1, Me.txtURL.Text, "(*)", vbBinaryCompare) < 1 Then

        MsgBox "输入URL中没有包含可替换部分:(*)"
        Exit Sub

    End If
    
    If Me.txtStart.Text = "" Or Not IsNumeric(Me.txtStart.Text & "") Then
        
        MsgBox "输入的开始数字不合法或为空"
        Exit Sub

    End If
    
    If Me.txtEnd.Text = "" Or Not IsNumeric(Me.txtEnd.Text & "") Then
        
        MsgBox "输入的结束数字不合法或为空"
        Exit Sub

    End If
    
    Dim strURL As String
    
    Dim i As Long
    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
        
        
    For i = Me.txtStart.Text To Me.txtEnd.Text
        
        Me.Caption = i
        
        strURL = Replace(Me.txtURL.Text, "(*)", i, 1, 1, vbBinaryCompare)
        
        iWeb.URL = strURL
        iWeb.CharSet = MainForm.cboCharSet.List(MainForm.cboCharSet.ListIndex)
        Dim strResult As String
        strResult = restoreCRLF(iWeb.StartGetHtml)
        
        MainForm.txtSrc.Text = MainForm.txtSrc.Text & vbCrLf & strResult
        
        

    Next
    
    Set iWeb = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
Call ReadConfig(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveConfig(Me)
End Sub
