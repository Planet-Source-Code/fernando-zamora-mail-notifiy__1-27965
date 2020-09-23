Attribute VB_Name = "modRegistry"
Option Explicit

'============================================
'   Registry - Load Settings
'============================================
Public Sub LoadAppSettings()

    ' Load User Settings
    frmSample.txtPassword = LeerIni(App.Path & "\config.ini", "SETTINGS", "PASSWORD", "")
    frmSample.txtUser = LeerIni(App.Path & "\config.ini", "SETTINGS", "USER", "")
    
    ' Load Mail Settings
    frmSample.txtDelay = LeerIni(App.Path & "\config.ini", "SETTINGS", "DELAY", "2")
    frmSample.txtServer = LeerIni(App.Path & "\config.ini", "SETTINGS", "SERVER", "")

    ' form load position
    frmSample.Left = LeerIni(App.Path & "\config.ini", "SETTINGS", "X", "0")
    frmSample.Top = LeerIni(App.Path & "\config.ini", "SETTINGS", "Y", "0")
End Sub


'============================================
'   Registry - Save Settings
'============================================
Public Sub SaveAppSettings()

    ' save User Settings
    If LeerIni(App.Path & "\config.ini", "SETTINGS", "PASSWORD", "") = "" Then
        GuardarIni App.Path & "\config.ini", "SETTINGS", "PASSWORD", Convert(frmSample.txtPassword)
    Else
        GuardarIni App.Path & "\config.ini", "SETTINGS", "PASSWORD", frmSample.txtPassword
    End If
    GuardarIni App.Path & "\config.ini", "SETTINGS", "USER", frmSample.txtUser
    
    ' save Mail Settings
    GuardarIni App.Path & "\config.ini", "SETTINGS", "DELAY", frmSample.txtDelay
    GuardarIni App.Path & "\config.ini", "SETTINGS", "SERVER", frmSample.txtServer

    ' save position
    GuardarIni App.Path & "\config.ini", "SETTINGS", "X", frmSample.Left
    GuardarIni App.Path & "\config.ini", "SETTINGS", "Y", frmSample.Top

End Sub


