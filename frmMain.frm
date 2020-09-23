VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Mail Setup"
   ClientHeight    =   2505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3810
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   2640
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1500
      Picture         =   "frmMain.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdNoRegistry 
      Caption         =   "Unsave"
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      ToolTipText     =   "Clears all fields and removes all saved information from system registry."
      Top             =   2130
      Width           =   1845
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2940
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Begin Monitoring!"
      Top             =   2130
      Width           =   795
   End
   Begin VB.Timer timMinutes 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3180
      Top             =   2910
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   720
      Picture         =   "frmMain.frx":114E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   2730
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1050
      TabIndex        =   2
      Top             =   1350
      Width           =   1665
   End
   Begin VB.PictureBox picMail 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   60
      Picture         =   "frmMain.frx":1A18
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   2670
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1050
      Width           =   1665
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   720
      Width           =   1665
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   1710
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2130
      Top             =   2910
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V1.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   15
      Top             =   2280
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   90
      Picture         =   "frmMain.frx":1E5A
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   4
      Left            =   900
      TabIndex        =   12
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   750
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Time Delay in Minutes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1470
      TabIndex        =   7
      Top             =   1740
      Width           =   1650
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "- Mail Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   1380
      Width           =   900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   840
      Top             =   0
      Width           =   4785
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      FillColor       =   &H00730009&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   0
      Top             =   0
      Width           =   825
   End
   Begin VB.Menu mnuTray 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuReadMailNow 
         Caption         =   "&Check Mail Now"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Quit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================
'   Mail Notification Application
'============================================
'============================================
'   I picked this up somewhere on
'   PlanetSourceCode, credit goes
'   to the guy who didn't put his
'   name in this code...I did some
'   major tweaking on it and viola,
'   a very simple mail notification.
'
'   Michael A.Schmidt
'   Written in October, 2000
'
'============================================
'   This application checks for new mail,
'   playing a .WAV file and showing the
'   number of new messages in your tray.
'============================================
'============================================
Public SoundByte As String          ' Path To WAV File
Public TrayText As String           ' Text For Tray Icon
Private MinutesElapsed As Integer   ' Minutes Elapsed Counter


'============================================
'   Form Load
'============================================
Private Sub Form_Load()

    SoundByte = App.Path & "\Mail.wav"

    ' Load Registry Settings
    LoadAppSettings
    
    ' Set Textboxes
    txtServer = AppSave.mServer
    txtUser = AppSave.mUser
    txtPassword = AppSave.mPassword
    txtDelay = AppSave.mDelay

    If txtServer.Text <> "" Then cmdGo_Click

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = 0 Then _
        If MsgBox("Quit Application?", vbYesNo, "Quit") = vbYes Then Cancel = -1
    
    If Cancel <> -1 Then
        DeleteIcon picMail(1)
        SaveAppSettings
    End If
        
End Sub




'============================================
'   Clear Button Click
'============================================
Private Sub cmdNoRegistry_Click()
    
    txtServer = ""      ' Clear Server
    txtUser = ""        ' Clear User
    txtPassword = ""    ' Clear Password
    DeleteAppSettings   ' Erase Registry Info

End Sub


'============================================
'   Go Button
'============================================
Private Sub cmdGo_Click()

    ' Save Application Settings
    AppSave.mServer = txtServer
    AppSave.mUser = txtUser
    AppSave.mPassword = txtPassword
    AppSave.mDelay = txtDelay
    SaveAppSettings
    
    ' Validate Fields
    Dim Control
    For Each Control In Controls
        If TypeOf Control Is TextBox Then
            If Len(Control.Text) = 0 Then
                MsgBox "Blank Data!", vbInformation
                Exit Sub
            End If
        End If
    Next

    Me.Hide
    timMinutes.Enabled = True

    ' Tray Icon Setup
    TrayText = "Mail Notification - " & AppSave.mServer
    AddIcon picMail(IconState.Idle), TrayText
    
    timMinutes_Timer


End Sub


'============================================
'   Menu - Quit
'============================================
Private Sub Mnu_Quit_Click()

    Unload Me

End Sub

'============================================
'   Go Button
'============================================
Private Sub mnuReadMailNow_Click()
    CheckNewMail
End Sub
'============================================
'   Menu - Setup
'============================================
Private Sub MnuSetup_Click()

    timMinutes.Enabled = False  ' Stop Counter
    Me.Show                     ' Show Form

End Sub


'============================================
'   Tray Pic - Mouse Move
'============================================
Private Sub picMail_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Select Case X
        Case trayLBUTTONUP
            ' Do Left Button Stuff
        Case trayRBUTTONUP
            PopupMenu mnuTray
        Case Else
    End Select
End Sub


'============================================
'   Timer - Minutes
'============================================
Private Sub timMinutes_Timer()
    '#############################
    '# Every minute this sub is
    '# called, we simply increment
    '# our counter or reset and
    '# check the mail.
    MinutesElapsed = MinutesElapsed + 1

    If MinutesElapsed = txtDelay Then
        CheckNewMail        ' Check Mail
        MinutesElapsed = 0  ' Reset Counter
    End If

End Sub


'============================================
'   Winsock - Data Arrival
'============================================
' Hmm...I left this sub alone...although it
' could use much improvement...
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    'Save the received data into strData variable
    Winsock1.GetData strData
            

    If Left$(strData, 1) = "+" Then
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                Winsock1.SendData "USER " & AppSave.mUser & vbCrLf
                
                'Here is the end of Winsock1_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & AppSave.mPassword & vbCrLf
                
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                Winsock1.SendData "STAT" & vbCrLf
                
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                TotalMails = CInt(Mid$(strData, 5, _
                              InStr(5, strData, " ") - 5))
                'If intMessages > 0 Then
                    '
                    'Oops. There is something in the mailbox!
                    'Change the session state
                    'm_State = POP3_RETR
                    '
                    'Increment the number of messages by one
                    'intCurrentMessage = intCurrentMessage + 1
                    '
                    'and we're sending to the server the RETR command in
                    'order to retrieve the first message
                    'Winsock1.SendData "RETR 1" & vbCrLf
                    
                'Else
                    'The mailbox is empty. Send the QUIT command to the
                    'server in order to close the session
                m_State = POP3_QUIT
                Winsock1.SendData "QUIT" & vbCrLf
                
                'MsgBox "You have not mail.", vbInformation
                'End If
            Case POP3_RETR
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                Winsock1.Close
                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                If TotalMails > 0 Then

                    MsgStatus = TotalMails & " New Message(s)!"
                    ChangeIcon frmMain.picMail(IconState.NewMail), MsgStatus
                    sndPlaySound SoundByte, 2

                Else
                    MsgStatus = "No New Mail - " & AppSave.mServer
                    ChangeIcon frmMain.picMail(IconState.Idle), MsgStatus
                End If
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            Winsock1.Close
            ChangeIcon frmMain.picMail(IconState.MailError), "POP3 ERROR: " & strData
    End If

End Sub


'============================================
'   Winsock - Error
'============================================
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ChangeIcon frmMain.picMail(IconState.MailError), Description
End Sub
