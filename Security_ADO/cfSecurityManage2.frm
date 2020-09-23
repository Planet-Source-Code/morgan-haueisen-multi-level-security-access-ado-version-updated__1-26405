VERSION 5.00
Begin VB.Form cfSecurityManage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   4695
   ClientTop       =   3210
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "cfSecurityManage2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   120
      TabIndex        =   15
      Top             =   4575
      Width           =   3585
      Begin VB.TextBox txtExpDays 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Text            =   "0"
         Top             =   165
         Width           =   495
      End
      Begin VB.CommandButton cmdSaveExpDays 
         Caption         =   "Save"
         Height          =   300
         Left            =   2730
         TabIndex        =   16
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Password Expiration Days: "
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   195
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   420
      Left            =   3795
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4665
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.Frame fraGroups 
      Caption         =   "&Group Membership"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1875
      Width           =   4695
      Begin VB.CommandButton cmdRemoveUserFromGroup 
         Caption         =   "<< &Remove"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddUserToGroup 
         Caption         =   "&Add >>"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.ListBox lstGroupsUserIn 
         Height          =   1815
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.ListBox lstGroupsUserNotIn 
         Height          =   1815
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblPasswordExp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password never expires "
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2745
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Member Of:"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Available Groups:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   1575
      End
   End
   Begin VB.Frame fraUsers 
      Caption         =   "&User Accounts"
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdEditUser 
         Caption         =   "&Edit User"
         Height          =   345
         Left            =   720
         TabIndex        =   25
         Top             =   1170
         Width           =   1815
      End
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Change &Password"
         Height          =   345
         Left            =   2850
         TabIndex        =   14
         Top             =   1170
         Width           =   1695
      End
      Begin VB.CommandButton cmdClearPassword 
         Caption         =   "&Clear Password"
         Height          =   345
         Left            =   2850
         TabIndex        =   5
         Top             =   780
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   780
         Width           =   855
      End
      Begin VB.CommandButton cmdNewUser 
         Caption         =   "&New..."
         Height          =   345
         Left            =   720
         TabIndex        =   3
         Top             =   780
         Width           =   855
      End
      Begin VB.ComboBox cboUsers 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "cfSecurityManage2.frx":000C
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox picEnterNew 
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   4830
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4830
      Begin VB.TextBox txtNewUser 
         Height          =   285
         Left            =   690
         MaxLength       =   20
         TabIndex        =   23
         Top             =   645
         Width           =   3570
      End
      Begin VB.CommandButton cmdNewUserOK 
         Caption         =   "Ok"
         Height          =   495
         Left            =   2580
         TabIndex        =   22
         Top             =   1560
         Width           =   1665
      End
      Begin VB.CommandButton cmdNewUserCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   795
         TabIndex        =   21
         Top             =   1560
         Width           =   1665
      End
      Begin VB.CheckBox chkNoExp 
         Caption         =   "User's password never expires."
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   1050
         Width           =   3465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Enter the new User's Login ID: "
         Height          =   195
         Left            =   615
         TabIndex        =   24
         Top             =   300
         Width           =   2250
      End
   End
End
Attribute VB_Name = "cfSecurityManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2000-2001

'/* This form will be used to manage user and group security.
Option Explicit

Dim goSecurity As clsSecurity

Dim msNewUser      As String
Dim msUserOldName  As String
Dim AddNew         As Boolean
                        
Private Function FillUserList() As Boolean
'/* This function will request a list of users from the
'/* server security object and place each user into the
'/* appropriate list (cboUsers).
'/* Returns:    True/False if ok.

On Error Resume Next
    
    Dim goSecurity  As clsSecurity
    Set goSecurity = New clsSecurity
    
    FillUserList = goSecurity.ListAllUsers(cboUsers, goApplication.SecurityDatabasePath, msNewUser)
    
    Set goSecurity = Nothing
End Function

Private Function FillGroupList(UserName As String) As Boolean
'/* This function will request an list of groups from the
'/* server security object and plac each group into the
'/* appropriate list object, depending on the current user's
'/* membership in each group.

'/* UserName - This string contains the user's name that will be used
'/*  to fill the group lists.

'/* Returns:    True/False if ok.

On Error Resume Next
    
    Dim goSecurity  As clsSecurity
    Set goSecurity = New clsSecurity
    
    FillGroupList = goSecurity.ListAllGroupsPerUser(lstGroupsUserNotIn, lstGroupsUserIn, UserName, goApplication.SecurityDatabasePath)
    
    Set goSecurity = Nothing
End Function

Private Sub cboUsers_Click()
    On Error Resume Next
    Call FillGroupList(cboUsers.Text)
    lblPasswordExp.Visible = CBool(cboUsers.ItemData(cboUsers.ListIndex))
End Sub

Private Sub cmdAddUserToGroup_Click()
    On Error Resume Next
    
    Dim goSecurity As clsSecurity
    Set goSecurity = New clsSecurity

    If cboUsers.ListIndex > -1 And lstGroupsUserNotIn.ListIndex > -1 Then
        '/* a user has been selected from the list.
        '/* AND a group has been selected to add.
        If goSecurity.AddUserToGroup(lstGroupsUserNotIn.Text, cboUsers.Text, goApplication.SecurityDatabasePath) Then
            Call FillGroupList(cboUsers.Text)
        End If
    End If
    
    Set goSecurity = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassword_Click()

'/* The password will be changed for the selected user.

    If lstGroupsUserIn.ListCount = 0 Then
        MsgBox "You must add user rights before changing the users password", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next

    Dim NewPassword As String
    Dim goSecurity As clsSecurity
    Set goSecurity = New clsSecurity

    If cboUsers.ListIndex > -1 Then
        NewPassword = InputBox("Enter a new Password for user " & cboUsers.Text, "Change a Password", "")
        NewPassword = LCase(Trim(NewPassword))
        If Len(NewPassword) < 5 Or Len(NewPassword) > 14 Then
            MsgBox "Password must be at between 5 and 14 characters long", vbExclamation
            Exit Sub
        End If
        If NewPassword > vbNullString Then
            '/* Change to the new Password */
            If goSecurity.ChangePassword(cboUsers.Text, NewPassword, NewPassword, "ChangePassword", goApplication.SecurityDatabasePath) Then
                MsgBox "Password has been changed.", vbInformation, Me.Caption
            Else
                MsgBox "Unable to change the user's password.", vbExclamation, Me.Caption
            End If
        End If
    End If
    Set goSecurity = Nothing

End Sub

Private Sub cmdClearPassword_Click()
'/* The password will be cleared out for the selected user.

    On Error Resume Next

    Dim goSecurity As clsSecurity

    Set goSecurity = New clsSecurity

    If cboUsers.ListIndex > -1 Then
        '/* a user account has been selected
        If goSecurity.ClearPassword(cboUsers.Text, goApplication.SecurityDatabasePath) Then
            MsgBox "Password has been reset to password.", vbInformation, Me.Caption
        Else
            MsgBox "Unable to reset the user's password.", vbExclamation, Me.Caption
        End If
    End If
    
    Set goSecurity = Nothing
End Sub

Private Sub cmdDeleteUser_Click()
'/* Will executed when the user selects
'/* the Delete command button.

    Dim bSelectComboItem As Boolean
    Dim goSecurity As clsSecurity

    Set goSecurity = New clsSecurity
    
    On Error Resume Next
    If cboUsers.Text = goUser.UserName Then
        MsgBox "You can not delete your own account", vbInformation
        Exit Sub
    End If
    
    If cboUsers.ListIndex > -1 Then
        If goSecurity.DeleteUser(cboUsers.Text, goApplication.SecurityDatabasePath) Then
            Call FillUserList
            '/* use list not itemdata
            cboUsers.ListIndex = SelectComboItem(cboUsers, goUser.UserName, False)
            cboUsers.Refresh
        End If
    End If
    Set goSecurity = Nothing

End Sub

Private Sub cmdEditUser_Click()
'/* Will be executed when the user selects
'/* the Edit User command button.  The picNewUser will become visible.

    On Error Resume Next
    
    AddNew = False
    picEnterNew.Visible = True
    picEnterNew.ZOrder
    
    chkNoExp.Value = IIf(cboUsers.ItemData(cboUsers.ListIndex), vbChecked, vbUnchecked)
    txtNewUser.Text = cboUsers.Text
    msUserOldName = cboUsers.Text
    txtNewUser.SetFocus

End Sub

Private Sub cmdNewUser_Click()
'/* Will be executed when the user selects
'/* the New command button.  The picNewUser will become visible.

    On Error Resume Next
    
    AddNew = True
    picEnterNew.Visible = True
    picEnterNew.ZOrder

    chkNoExp.Value = vbUnchecked
    txtNewUser = vbNullString
    msUserOldName = vbNullString
    txtNewUser.SetFocus
End Sub

Private Function SelectComboItem(Cbo As ComboBox, SearchValue As Variant, SearchItemData As Boolean) As Long
 '/* This function will pick an item in the combo box if it exists.
 '/* It will be smart enough to know whether to look for the item in
 '/* the ItemData() or the List().
 Dim i       As Integer
 Dim itemp   As Integer
    
    itemp = -1
    If Len(SearchValue) <> 0 Then
        If SearchItemData Then
            '/* Search thru the item data
            For i = 0 To Cbo.ListCount - 1
                If Cbo.ItemData(i) = SearchValue Then
                    itemp = i
                    Exit For
                End If
            Next i
            
        Else
            '/* Search thru the list values
            For i = 0 To Cbo.ListCount - 1
                If Trim(Cbo.List(i)) = SearchValue Then
                    itemp = i
                    Exit For
                End If
            Next i
        End If
    End If
    SelectComboItem = itemp

End Function

Private Sub cmdNewUserCancel_Click()
    picEnterNew.Visible = False
End Sub

Private Sub cmdNewUserOK_Click()
'/* Will be executed when the user selects
'/* the New command button.
'/* The new user's
'/* name will be inserted into cboUsers and the group lists
'/* will reflect the new user's membership in groups.

    On Error Resume Next
    
    picEnterNew.Visible = False
    msNewUser = LCase(Trim(txtNewUser))
    If msUserOldName = vbNullString Then msUserOldName = msNewUser

    '/* the frmNewUser should have set the me.NewUser property.
    If Len(msNewUser) > 0 Then
        If AddNew Then
            If Not goSecurity.CreateUser(msNewUser, chkNoExp.Value, goApplication.SecurityDatabasePath, goApplication.SystemID) Then
                MsgBox "Unable to create new user", vbExclamation
            End If
        Else
            If Not goSecurity.EditUser(msUserOldName, msNewUser, chkNoExp.Value, goApplication.SecurityDatabasePath, goApplication.SystemID) Then
                MsgBox "Unable to edit user", vbExclamation
            End If
        End If
        Call FillUserList
    End If

End Sub

Private Sub cmdRemoveUserFromGroup_Click()
'/* This procedure will be executed when when the user selects
'/* the <<<Remove command button.  The group the user has highlighted
'/* in the Member of list will be moved to the Available Groups list.

    On Error Resume Next
    
    Dim goSecurity As clsSecurity
    Set goSecurity = New clsSecurity
    
    If cboUsers.ListIndex > -1 And lstGroupsUserIn.ListIndex > -1 Then
        '/* a user has been selected from the list.
        '/* AND a group has been selected to remove.
        If goSecurity.RemoveUserFromGroup(lstGroupsUserIn.Text, cboUsers.Text, goApplication.SecurityDatabasePath) Then
            '/* refill the group list to show the changes.
            Call FillGroupList(cboUsers.Text)
        End If
    End If
    Set goSecurity = Nothing
End Sub

Private Sub cmdSaveExpDays_Click()
    Call goSecurity.ExpDaysSet(txtExpDays, goApplication.SecurityDatabasePath)
End Sub

Private Sub Form_Load()
'/* This procedure will fill the cboUsers list from the current
'/* security file.

    'cScreen.CenterForm Me
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Me.Caption = App.Title & " - Manage Security"

    On Error Resume Next
    
    Set goSecurity = New clsSecurity
    
    Call FillUserList
    If cboUsers.ListCount > 0 Then
        cboUsers.ListIndex = 0
        Call FillGroupList(cboUsers.Text)
    End If
    
    txtExpDays = goSecurity.ExpDaysGet(goApplication.SecurityDatabasePath)
    
    Screen.MousePointer = vbDefault
    
    'If cTile.MaxColors(Me) > 300 Then
    '   Call cTile.TileBackground(Me, frmArt!Image3, 0)
    '   Call cTile.Shadow(Me, fraUsers, 60, -0.3)
    '   Call cTile.Shadow(Me, fraGroups, 60, -0.3)
    'End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set goSecurity = Nothing
    Set cfSecurityManage = Nothing
End Sub


Private Sub txtExpDays_GotFocus()
    txtExpDays.SelStart = 0
    txtExpDays.SelLength = Len(txtExpDays)
End Sub


Private Sub txtExpDays_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 48 To 57, 8
        '/* is a number or backspace
    Case Else
        KeyAscii = 0
    End Select
End Sub


Private Sub txtNewUser_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 48 To 57, 97 To 122, 8
    Case Else
        KeyAscii = False
    End Select

End Sub


