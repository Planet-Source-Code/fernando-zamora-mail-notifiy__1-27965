VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   0  'None
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4275
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSample.frx":08CA
   ScaleHeight     =   2805
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MailNotify.VBMail MyBox 
      Left            =   3735
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer timMinutes 
      Interval        =   60000
      Left            =   3735
      Top             =   1530
   End
   Begin VB.PictureBox picMail 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   450
      Picture         =   "frmSample.frx":7DA8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   990
      Picture         =   "frmSample.frx":8672
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1530
      Picture         =   "frmSample.frx":8AB4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2070
      Picture         =   "frmSample.frx":937E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   375
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmSample.frx":9C48
      Top             =   2010
      Width           =   2655
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   345
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   1635
      Width           =   510
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   345
      TabIndex        =   0
      Top             =   630
      Width           =   1755
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      TabIndex        =   2
      Top             =   1275
      Width           =   1800
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2175
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1710
   End
   Begin VB.CommandButton cmdCheckMailbox 
      Caption         =   "Check"
      Height          =   360
      Left            =   3075
      TabIndex        =   4
      Top             =   2010
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   3645
      Top             =   180
      Width           =   375
   End
   Begin VB.Menu mnuMailMenu 
      Caption         =   "MailMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckMail 
         Caption         =   "Check Mail"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Fernando Zamora
'  The Red Devil Co.
'  This is an enhacement of a source i found in PSC
'  But i couldn´t find the name of original autor
'  so if he read this,i'll like to give thanks for this piece of code
'  i have add Skins,encoded password and get the last mail
'  Best regards from Spain

Option Explicit

Dim MailTray As New clsTray     ' - Mail Tray
Dim MinutesElapsed As Integer   ' - Mail Check Timer


'================================
'   Form Load
'================================
Private Sub Form_Load()
    'RoundedForm
    Dim hrgn As Long
    hrgn = CreateRoundRectRgn(0, 0, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels), 50, 50)
    SetWindowRgn hWnd, hrgn, True
    DeleteObject hrgn
    
    
    MailTray.ShowIcon Me
    MailTray.ChangeIcon Me, picMail.Item(0)
    
    LoadAppSettings
    
    If Len(txtServer) <> 0 Then
        Me.Hide
        cmdCheckMailbox_Click
        'MYBox_NewMail (74)
    End If

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
formdrag Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1.Visible = True
End Sub

'================================
'   Form Resize
'================================
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Me.Hide

End Sub


'================================
'   Form Unload
'================================
Private Sub Form_Unload(Cancel As Integer)

    MailTray.RemoveIcon Me
    SaveAppSettings
    End
End Sub


'================================
'   Check Mailbox Button
'================================
Private Sub cmdCheckMailbox_Click()

    MyBox.CheckNewMail
    MailTray.ChangeIcon Me, picMail.Item(1)
    MailTray.ChangeToolTip Me, "Checking Mail (" & MyBox.Server & ")"

End Sub


Private Sub Image1_Click()
frmSample.WindowState = 1
End Sub

'================================
'   Menu Check Mail
'================================
Private Sub mnuCheckMail_Click()
    cmdCheckMailbox_Click
End Sub


'================================
'   Menu Setup
'================================
Private Sub mnuSetup_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'================================
'   OBJECT Event New Mail
'================================
Private Sub MYBox_NewMail(NumMsgs As Integer)
    Static OldNumMsgs As Integer
    MailTray.ChangeIcon Me, picMail.Item(2)
    MailTray.ChangeToolTip Me, NumMsgs & " New Message(s)on " & txtServer & "!"
    If NumMsgs > OldNumMsgs Then
    Form2.Label1.Caption = ParseHeader(MyBox.Lastmesagge)
    Form2.Show
    Form2.Left = Screen.Width
    Form2.Top = Screen.Height - 1250
    'Form2.Refresh
    Dim i As Long
    For i = 3000 To 0 Step -1
        Form2.Left = Form2.Left - 1
        'DoEvents
        Next
    End If
    
    OldNumMsgs = NumMsgs
End Sub


'================================
'   OBJECT Event Noisy
'================================
Private Sub MYBox_Noisy(POPresponse As String)
    txtLog = POPresponse & vbCrLf & txtLog
End Sub


'================================
'   OBJECT Event No Mail
'================================
Private Sub MYBox_NoMail()
    MailTray.ChangeIcon Me, picMail.Item(0)
    MailTray.ChangeToolTip Me, "No New Mail"
End Sub


'================================
'   OBJECT Event Error
'================================
Private Sub MYBox_SockError(ErrorStats As String)
    MailTray.ChangeIcon Me, picMail.Item(3)
    MailTray.ChangeToolTip Me, ErrorStats
End Sub


'================================
'   Timer
'================================
Private Sub timMinutes_Timer()
    '#############################
    '# Every minute this sub is
    '# called, we simply increment
    '# our counter or reset and
    '# check the mail.
    MinutesElapsed = MinutesElapsed + 1

    If MinutesElapsed = txtDelay Then
        cmdCheckMailbox_Click ' Check Mail
        MinutesElapsed = 0    ' Reset Counter
    End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remember..... The value of X will be different if the icon is minimised
' to the system tray.  The values in this case will be as follows,
'       7680   ' MouseMove
'       7695   ' Left MouseDown
'       7710   ' Left MouseUp
'       7725   ' Left DoubleClick
'       7740   ' Right MouseDown
'       7755   ' Right MouseUp
'       7770   ' Right DoubleClick
If MailTray.bRunningInTray Then          'Check to see if form is in the system tray
    Select Case X                           'If it is, use X to get message value
        Case 7755: PopupMenu Me.mnuMailMenu, vbPopupMenuRightButton
        Case 7725: Me.Show: Me.WindowState = vbNormal
    End Select
End If

End Sub


Private Sub txtPassword_Change()
    'se debe encriptar la clave para que no se vea al
    'tenerla guardada ni con un unmask
    MyBox.Password = Convert(txtPassword)
End Sub
Private Sub txtServer_Change()
    MyBox.Server = txtServer
End Sub
Private Sub txtUser_Change()
    MyBox.User = txtUser
End Sub

Private Function ParseHeader(Header As String) As String
    Dim HeaderStart As Long
    Dim HeaderEnd As Long
    HeaderStart = InStr(1, Header, "<")
    HeaderEnd = InStr(HeaderStart, Header, ">")
    ParseHeader = Mid$(Header, HeaderStart, HeaderEnd)
End Function
