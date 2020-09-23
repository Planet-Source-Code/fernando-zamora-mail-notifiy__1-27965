VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   840
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1125
      Top             =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Mail !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   405
      TabIndex        =   1
      Top             =   135
      Width           =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW MAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   450
      Width           =   2520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Dim i As Long
For i = 3000 To 0 Step -1
    Form2.Left = Form2.Left + 1
    DoEvents
Next
Unload Me
End Sub

Private Sub Label1_Click()
Timer1.Enabled = False
Dim i As Long
For i = 3000 To 0 Step -1
    Form2.Left = Form2.Left + 1
    DoEvents
Next
Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    For i = 3000 To 0 Step -1
        Form2.Left = Form2.Left + 1
        DoEvents
        Next
Unload Me
End Sub
