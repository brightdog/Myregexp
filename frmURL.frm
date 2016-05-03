VERSION 5.00
Begin VB.Form frmURL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入URL"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   Icon            =   "frmURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtReferer 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   5640
      Width           =   7815
   End
   Begin VB.TextBox txtPost 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   900
      Width           =   8895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4980
      TabIndex        =   2
      Top             =   6240
      Width           =   1785
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1980
      TabIndex        =   1
      Top             =   6240
      Width           =   1785
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   8505
   End
   Begin VB.Label lblReferer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referer"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   5700
      Width           =   750
   End
   Begin VB.Label lblU_R_ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmURL.frx":000C
      Height          =   540
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "frmURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public btnOK As Boolean
Public btnCancel As Boolean
Public lastURL As String
Public lastPost As String

Private Sub cmdCancel_Click()
    btnCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
        Call SaveConfig(Me)
    btnOK = True
    Me.Hide
End Sub

Public Property Get URL() As String

    URL = Me.txtURL.Text

End Property
Public Property Get POST() As String

    POST = Me.txtPost.Text

End Property


Private Sub Form_Load()
    Me.txtURL.Text = lastURL
    Me.txtPost.Text = lastPost
    Me.txtURL.SelStart = 0
    Me.txtURL.SelLength = Len(Me.txtURL.Text)
    Call ReadConfig(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveConfig(Me)
End Sub
