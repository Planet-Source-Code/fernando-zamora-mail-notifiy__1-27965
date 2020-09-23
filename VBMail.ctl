VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl VBMail 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VBMail.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   420
   ToolboxBitmap   =   "VBMail.ctx":0972
   Begin MSWinsockLib.Winsock Socket 
      Left            =   780
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "VBMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   THIS CODE IS COPYRIGHT (C) 2000 MICHAEL A. SCHMIDT
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   Author: Michael A. Schmidt
'   Email:  mschmidt@mtdmarketing.com
'   Date:   January 12, 2001
'   Referances: MSWINSCK.OCX
'   Posted: Planet Source Code - http://www.planetsourcecode.com/vb
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'   Description:
'       This Class allows you to check your mail through an OBJECT.
'       Create multiple objects, check your mail through them, =)
'           You may use this inside your project or through the ActiveX DLL.
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private iState As Integer
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event NewMail(NumMsgs As Integer)
Public Event NoMail()
Public Event Noisy(POPresponse As String)
                        ' If you wish to view server reponse for debugging purposes
                        ' in your app, then this event will return all server reponse.
Public Event SockError(ErrorStats As String)
                        ' When winsock/pop generates an error, we simply pass it
                        ' on to the controller, by raising the event and passing.
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarUser As String
Private mvarPassword As String
Private mvarMailPort As Integer
Private mvarServer As String
Private strLastmessage As String
Private iMessages As Integer
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
'   Check New Mail
'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
Public Sub CheckNewMail()
     If Server = "" Then Exit Sub

    ' Clear Mail Stage
    iState = 0

    ' Connect To Mail Server
    Socket.Close
    Socket.LocalPort = 0
    Socket.Connect Server, MailPort
    ' Mail Port Default = 110
    
    RaiseEvent Noisy("Connecting (" & Server & ")")

End Sub


'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
'   Socket Error ( POP Error )
'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim ErrorStats As String

    ErrorStats = Number & " : " & Description
    RaiseEvent SockError(ErrorStats)

End Sub


'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
'   Data Arrival ( Heart 'O Program )
'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
'   There are two ways to write this. I have chosen the
'   first method.
'
'   First, simply count the reponses you have received
'   and assume to know the protocol. This is not the
'   better of the two, but it is the simpler, requiring
'   little knowledge of parsing and protocol.
'
'   Second, you could translate every response to tell
'   what stage you are at. This is accurate, but may not
'   apply to the remote server's protocol.
'/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|/-\|
Dim iPacket As String
Dim iCom As String

On Local Error GoTo BadSub
    Socket.GetData iPacket      ' - Pull Server Packet
    iCom = Left(iPacket, 1)     ' - Grab 1st Character
    If iState = 4 Then
       strLastmessage = ParseHeader(iPacket)
    Else
        RaiseEvent Noisy(iPacket)   ' - Pass Reponse Via Event
    End If
    If iCom = "+" Then GoTo GoodSub ' - If Good, Process Mail
    If iCom = "-" Then GoTo BadSub  ' - If Bad, Process Error

Exit Sub
GoodSub:

    

    Select Case iState
    Case 0: Socket.SendData "USER " & User & vbCrLf        ' - USER INFORMATION
            iMessages = 0
            RaiseEvent Noisy("Sending USER...")
    Case 1: Socket.SendData "PASS " & Password & vbCrLf    ' - PASS INFORMATION
            RaiseEvent Noisy("Sending PASS...")
    Case 2: Socket.SendData "STAT" & vbCrLf                ' - GETMAIL COMMAND
            RaiseEvent Noisy("Checking Mail...")
    Case 3: iMessages = CInt(Mid$(iPacket, 5, InStr(5, iPacket, " ") - 5))
            If iMessages > 0 Then Socket.SendData "top " & iMessages & " 0" & vbCrLf
    Case 4:
            Socket.SendData "QUIT" & vbCrLf
            If iMessages = 0 Then RaiseEvent NoMail
            If iMessages > 0 Then RaiseEvent NewMail(iMessages)
    Case 5: Socket.Close                                   ' - CLOSE COMMAND
            RaiseEvent Noisy("Socket Closed...")
    Case Else:
    End Select

    ' Next Stage of Mail
    iState = iState + 1

Exit Sub
BadSub:
    RaiseEvent SockError("POP Error: " & iPacket)
    Debug.Print ("POP Error: " & iPacket)

End Sub


Public Property Let Server(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    mvarServer = vData
End Property


Public Property Get Server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    Server = mvarServer
End Property



Public Property Let MailPort(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MailPort = 5
    mvarMailPort = vData
End Property


Public Property Get MailPort() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MailPort
    If mvarMailPort = 0 Then mvarMailPort = 110
    MailPort = mvarMailPort
End Property



Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let User(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.User = 5
    mvarUser = vData
End Property


Public Property Get User() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.User
    User = mvarUser
End Property

Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 420
End Sub

Public Property Get Lastmesagge() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.User
    Lastmesagge = strLastmessage
End Property

Private Function ParseHeader(Header As String) As String
    Dim HeaderStart As Long
    Dim HeaderEnd As Long
    HeaderStart = InStr(1, Header, "<")
    HeaderEnd = InStr(HeaderStart, Header, ">")
    ParseHeader = Mid$(Header, HeaderStart, HeaderEnd)
End Function
