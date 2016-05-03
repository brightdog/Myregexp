VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPop 
   Caption         =   "Resault"
   ClientHeight    =   6330
   ClientLeft      =   2730
   ClientTop       =   3240
   ClientWidth     =   8190
   Icon            =   "frmPop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtResault 
      Height          =   6285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   11086
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmPop.frx":00D2
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    Me.txtResault.width = Me.width - 100
    Me.txtResault.Height = Me.Height - 400
End Sub

Private Sub txtResault_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
