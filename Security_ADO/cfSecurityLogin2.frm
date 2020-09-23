VERSION 5.00
Begin VB.Form cfSecurityLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   3735
   ClientTop       =   2715
   ClientWidth     =   4230
   Icon            =   "cfSecurityLogin2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   4230
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2520
      Width           =   4230
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change Password"
         Height          =   480
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   480
         Left            =   2880
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   480
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPassWd 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1425
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1215
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtPassWd 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1425
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1815
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3420
      Picture         =   "cfSecurityLogin2.frx":0442
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblChgPW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Re-enter to verify, then click OK."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   645
      TabIndex        =   8
      Top             =   1575
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblChgPW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter NEW password ..."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   645
      TabIndex        =   7
      Top             =   975
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblUserID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "cfSecurityLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2000-2001

Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
    
'Private cPlay As New clsPlaySound

Dim ChangeFlag As Boolean
Dim PassTry As Integer

Private Sub cmdCancel_Click()
    goUser.UserName = vbNullString
    PassTry = 0
    Unload Me
End Sub

Private Sub cmdChange_Click()
    ChangeFlag = Not ChangeFlag
    If Not ChangeFlag Then
        Me.Height = 1980
        cmdChange.Caption = "Change Password"
        lblChgPW(1).Visible = False
        lblChgPW(2).Visible = False
        txtPassWd(1).Visible = False
        txtPassWd(2).Visible = False
    Else
        cmdChange.Caption = "Cancel Change"
        Me.Height = 3375
        lblChgPW(1).Visible = True
        lblChgPW(2).Visible = True
        txtPassWd(1).Visible = True
        txtPassWd(2).Visible = True
        txtPassword.SetFocus
    End If
End Sub

Private Sub cmdOk_Click()
'/* This procedure will submit the parameters entered by the user on the login
'/* form to the Login method of the Security object.
'/* It will set the SuccessfulLogin property
'/* of the User class to True if the login was successful.

    Dim goSecurity As clsSecurity
    
    Dim tNewPassWord1 As String
    Dim tNewPassWord2 As String
    Dim tUserName     As String
    Dim tPassWord     As String
    Dim bOk           As Integer
    Dim tString       As String
    
    On Error GoTo HandleErrors
    
    Me.MousePointer = vbHourglass
    
    Set goSecurity = New clsSecurity
    
    tUserName = txtUserID
    tPassWord = Trim(txtPassword)
    tNewPassWord1 = Trim(txtPassWd(1))
    tNewPassWord2 = Trim(txtPassWd(2))
    
    If ChangeFlag Then
        If Len(txtPassWd(1)) > 0 And Len(txtPassWd(1)) < 5 Then
            MsgBox "Password must be at between 5 and 15 characters long", vbInformation
            Set goSecurity = Nothing
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf txtPassWd(1) <> txtPassWd(2) Then
            MsgBox "New password does not match the verify password.", vbInformation
            Set goSecurity = Nothing
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If goSecurity.ChangePassword(tUserName, tNewPassWord1, tNewPassWord2, tPassWord, goApplication.SecurityDatabasePath) Then
            MsgBox "Password has been changed.", vbInformation, Me.Caption
            tPassWord = tNewPassWord1
        Else
            MsgBox "Unable to change the user's password.", vbExclamation, Me.Caption
        End If
    End If

    
    '/* Check for valid user ID and password
    bOk = goSecurity.Login(tUserName, tPassWord, goApplication.SecurityDatabasePath, goApplication.SystemID)
    Select Case bOk
    Case Is = 0
        If PassTry > 1 Then
            'cPlay.PlaySoundResource 1002
            MsgBox "Please do not attempt to use this program" & vbLf & "without the correct password.", vbCritical, "Ending Program"
            Set goSecurity = Nothing
            End
        Else
            'cPlay.PlaySoundResource 1001, , True
            MsgBox "Incorrect password - please re-enter!" & vbLf & "Your password is case sensitive", vbInformation, "Password Entry Error"
            Set goSecurity = Nothing
            PassTry = PassTry + 1
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    Case Is > 0, -2 '/* Password has expired or is the word password
        tString = "You must select a new password before you can continue."
        If bOk = -2 Then '/* User's password=password
            txtPassword = "password"
            tPassWord = "password"
        Else '/* User's password has expired
            tString = "Your password is " & CStr(bOk) & " days old and has expired." & vbCrLf & tString
        End If
        MsgBox tString, vbExclamation
        ChangeFlag = True
        cmdChange.Caption = "Cancel Change"
        Me.Height = 3375
        lblChgPW(1).Visible = True
        lblChgPW(2).Visible = True
        txtPassWd(1).Visible = True
        txtPassWd(2).Visible = True
        txtPassWd(1).SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    Case Is < 0
        '/* Login Ok
        'cPlay.PlaySoundResource 1000
        goUser.UserName = tUserName
        goUser.Password = tPassWord
        PassTry = 0
    End Select
    
Exit_Procedure:
    Set goSecurity = Nothing
    Me.MousePointer = vbDefault
    ChangeFlag = False
    Unload Me
Exit Sub


HandleErrors:
    MsgBox Err.Number & vbCrLf & "cfSecurityLogin" & vbCrLf & Err.Description
    Resume Exit_Procedure
End Sub

Private Function sUserName() As String
Dim strBuffer As String * 255
Dim lngBufferLength As Long
Dim lngRet As Long
Dim strTemp As String, i As Integer

    lngBufferLength = 255
    lngRet = GetUserName(strBuffer, lngBufferLength)
    strTemp = Trim$(strBuffer)
    For i = 1 To Len(strTemp)
        If Asc(Mid$(strTemp, i, 1)) > 122 Or Asc(Mid$(strTemp, i, 1)) < 15 Then
            strTemp = Left$(strTemp, i - 1)
            Exit For
        End If
    Next i
    
    sUserName = Trim(LCase$(strTemp))
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    
    Me.Height = 1980
    Me.Caption = App.Title & " - LogIn"
    
    '/* SEE FORM_RESIZE FOR CENTERING */
    '/*********************************/
    
    txtUserID = sUserName
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    'cScreen.CenterForm Me
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    'If cTile.MaxColors(Me) > 300 Then
    '   Call cTile.TileBackground(Me, frmArt!Image1, 0)
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cfSecurityLogin = Nothing
End Sub

Private Sub txtPassWd_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 48 To 57, 97 To 122, 8
    Case 13
        KeyAscii = 0
        If index = 2 Then
            Call cmdOk_Click
        Else
            SendKeys "{TAB}"
        End If
    Case Else
        KeyAscii = False
    End Select
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 48 To 57, 97 To 122, 8
    Case 13
        KeyAscii = 0
        If ChangeFlag Then
            txtPassWd(1).SetFocus
        Else
            Call cmdOk_Click
        End If
    Case Else
        KeyAscii = False
    End Select
End Sub


Private Sub txtUserID_GotFocus()
    txtUserID.SelStart = 0
    txtUserID.SelLength = Len(txtUserID)
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 48 To 57, 97 To 122, 8
    Case 13
        KeyAscii = 0
        SendKeys "{TAB}"
    Case Else
        KeyAscii = False
    End Select
End Sub


