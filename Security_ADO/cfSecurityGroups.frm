VERSION 5.00
Begin VB.Form cfSecurityGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Group Names"
   ClientHeight    =   5265
   ClientLeft      =   2895
   ClientTop       =   1665
   ClientWidth     =   7365
   Icon            =   "cfSecurityGroups.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   720
      Left            =   5055
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.Frame Frame4 
      Caption         =   "System ID"
      Height          =   1320
      Left            =   60
      TabIndex        =   5
      Top             =   3780
      Width           =   7155
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   405
         Left            =   585
         TabIndex        =   7
         Top             =   780
         Width           =   870
      End
      Begin VB.TextBox txtSystemID 
         Height          =   285
         Left            =   165
         MaxLength       =   42
         TabIndex        =   6
         Top             =   390
         Width           =   6690
      End
      Begin VB.Label Label2 
         Caption         =   "WARNING: If you change the SystemID, you will need to add the new System ID to the InitSecurity call."
         Height          =   390
         Left            =   1710
         TabIndex        =   8
         Top             =   780
         Width           =   5010
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Left            =   4215
      TabIndex        =   1
      Top             =   2310
      Width           =   3000
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   405
         Left            =   1890
         TabIndex        =   3
         Top             =   705
         Width           =   870
      End
      Begin VB.TextBox txtExpDays 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2220
         TabIndex        =   2
         Text            =   "0"
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Password Expiration Days:"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Names"
      Height          =   3510
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   4080
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   405
         ScaleHeight     =   2505
         ScaleWidth      =   3555
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   915
         Width           =   3555
         Begin VB.ListBox cboGroups 
            Height          =   2400
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   2250
         End
         Begin VB.CommandButton cmdDeleteUser 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2520
            TabIndex        =   15
            Top             =   975
            Width           =   855
         End
         Begin VB.CommandButton cmdNewUser 
            Caption         =   "&New..."
            Height          =   345
            Left            =   2520
            TabIndex        =   14
            Top             =   45
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2520
            TabIndex        =   13
            Top             =   510
            Width           =   855
         End
      End
      Begin VB.TextBox txtGroupName 
         Height          =   315
         Left            =   405
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   420
         Width           =   2250
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2925
         TabIndex        =   10
         Top             =   405
         Width           =   855
      End
   End
End
Attribute VB_Name = "cfSecurityGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2000-2001

Option Explicit
Dim oSecurity As New clsSecurity
Dim AddNew As Boolean

Private Sub cboGroups_Click()
    txtGroupName = cboGroups
    If cboGroups.ListCount > -1 Then
        cmdEdit.Enabled = True
        cmdDeleteUser.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeleteUser_Click()
    cmdEdit.Enabled = False
    cmdDeleteUser.Enabled = False
    
    If oSecurity.DeleteGroup(cboGroups, goApplication.SecurityDatabasePath) Then
        oSecurity.ListAllGroups cboGroups, goApplication.SecurityDatabasePath
        txtGroupName = cboGroups
    Else
        MsgBox "Unable to delete Group", vbCritical
    End If
End Sub

Private Sub cmdEdit_Click()
    txtGroupName.Locked = False
    cmdSave.Enabled = True
    
    oSecurity.CreateGroup txtGroupName, goApplication.SecurityDatabasePath, cboGroups
    
    Picture1.Enabled = False
    AddNew = False
    txtGroupName.SetFocus
End Sub

Private Sub cmdNewUser_Click()
    txtGroupName = vbNullString
    txtGroupName.Locked = False
    cmdSave.Enabled = True
    Picture1.Enabled = False
    AddNew = True
    txtGroupName.SetFocus
End Sub


Private Sub cmdSave_Click()
  Dim tGroupName As String
    
    tGroupName = Trim(txtGroupName)
    If tGroupName > vbNullString Then
        If AddNew Then
            oSecurity.CreateGroup tGroupName, goApplication.SecurityDatabasePath
        Else
            oSecurity.CreateGroup tGroupName, goApplication.SecurityDatabasePath, cboGroups
        End If
    End If
    
    oSecurity.ListAllGroups cboGroups, goApplication.SecurityDatabasePath, tGroupName
    
    txtGroupName.Locked = True
    cmdSave.Enabled = False
    
    Picture1.Enabled = True
    cmdEdit.Enabled = False
    cmdDeleteUser.Enabled = False

    AddNew = False
End Sub


Private Sub Command1_Click()
    oSecurity.ExpDaysSet txtExpDays, goApplication.SecurityDatabasePath
End Sub

Private Sub Command2_Click()
    oSecurity.SystemIDSet txtSystemID, goApplication.SecurityDatabasePath
End Sub


Private Sub Form_Load()
    'cScreen.CenterForm Me
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Me.Caption = App.Title & " - Manage Groups"
    oSecurity.ListAllGroups cboGroups, goApplication.SecurityDatabasePath
    txtExpDays = oSecurity.ExpDaysGet(goApplication.SecurityDatabasePath)
    txtSystemID = oSecurity.SystemIDGet(goApplication.SecurityDatabasePath)
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oSecurity = Nothing
    Set cfSecurityGroups = Nothing
End Sub


Private Sub txtGroupName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdSave_Click
    End If
End Sub


