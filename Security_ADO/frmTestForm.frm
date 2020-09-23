VERSION 5.00
Begin VB.Form frmTestForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmTestForm"
   ClientHeight    =   4965
   ClientLeft      =   3270
   ClientTop       =   2580
   ClientWidth     =   6315
   Icon            =   "frmTestForm.frx":0000
   LinkTopic       =   "FormMenuEnable"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log In"
      Height          =   435
      Left            =   195
      TabIndex        =   9
      Top             =   735
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   90
      TabIndex        =   6
      Top             =   2850
      Width           =   2475
      Begin VB.CommandButton cmdGroups 
         Caption         =   "Manage Groups"
         Height          =   435
         Left            =   165
         TabIndex        =   7
         Top             =   225
         Width           =   2025
      End
      Begin VB.Label Label2 
         Caption         =   "Used for first time setting of access groups.  Not to be included in the final application"
         Height          =   825
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   690
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1305
      Left            =   2910
      TabIndex        =   3
      Top             =   2430
      Width           =   3285
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open WorkingBD.mdb"
         Height          =   435
         Left            =   345
         TabIndex        =   4
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label Label2 
         Caption         =   "Open a password protected database. (PWD=123456)"
         Height          =   480
         Index           =   1
         Left            =   345
         TabIndex        =   5
         Top             =   660
         Width           =   2520
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Left            =   90
      TabIndex        =   0
      Top             =   1470
      Width           =   2460
      Begin VB.CommandButton cmdSecurity 
         Caption         =   "Manage Security"
         Height          =   435
         Left            =   225
         TabIndex        =   1
         Top             =   195
         Width           =   2025
      End
      Begin VB.Label Label5 
         Caption         =   "Add/remove users and set user access levels."
         Height          =   570
         Left            =   150
         TabIndex        =   2
         Top             =   675
         Width           =   2085
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Try Logging in as"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2010
      Left            =   2910
      TabIndex        =   13
      Top             =   90
      Width           =   3240
   End
   Begin VB.Label lblLogIn 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LogIn ID: "
      Height          =   285
      Left            =   165
      TabIndex        =   12
      Top             =   285
      Width           =   2070
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTestForm.frx":000C
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   2910
      TabIndex        =   11
      Top             =   3870
      Width           =   3285
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "User mhaueis has adminstrator access"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2910
      TabIndex        =   10
      Top             =   2100
      Width           =   3240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open a secure DB"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuInput 
      Caption         =   "Input"
      Begin VB.Menu mnuProduction 
         Caption         =   "Production"
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "Plan"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuFinancial 
         Caption         =   "Financial"
      End
      Begin VB.Menu mnuGeneral 
         Caption         =   "General"
      End
   End
   Begin VB.Menu mnuTables 
      Caption         =   "Tables"
      Begin VB.Menu mnuTbl1 
         Caption         =   "Table1"
      End
      Begin VB.Menu mnuTbl2 
         Caption         =   "Tabel2"
      End
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "Security"
      Begin VB.Menu mnuManageSecurity 
         Caption         =   "Manage Security"
      End
      Begin VB.Menu mnuGroups 
         Caption         =   "Manage Groups"
      End
   End
End
Attribute VB_Name = "frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGroups_Click()
    Call mnuGroups_Click
End Sub


Private Sub cmdOpen_Click()
    Call mnuOpen_Click
End Sub

Private Sub cmdSecurity_Click()
    Call mnuManageSecurity_Click
End Sub

Private Sub cmdLogIn_Click()
    '/* Clear current Log-in
    goUser.UserName = vbNullString
    '/* Load Log-in form
    cfSecurityLogin.Show vbModal
    '/* Set menu choices based on user's rights
    Call SetMenuAccess
    '/* Display Log-in ID
    lblLogIn = "LogIn ID: " & goUser.UserName
End Sub

Private Sub Form_Load()
    'cScreen.CenterForm Me
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 3
    
    '/* initilize application security
    Call InitSecurity(App.Path & "\WorkingDB.mdb", App.Path & "\Security.pwd")
    Call SetMenuAccess
    
    Label1 = " Try Logging in as" & vbCrLf & _
             "  Login-ID    Password" & vbCrLf & _
             "  --------------------" & vbCrLf & _
             "  bbutler     happy" & vbCrLf & _
             "  mjohnso     mjohnso" & vbCrLf & _
             "  asilva      password" & vbCrLf & _
             "  mhaueis     haueisen"
             
End Sub



Private Sub SetMenuAccess()
  Dim oSecurity As clsSecurity
  Dim vData() As Variant
  Dim i As Integer
    
    '/* Disable Secure Menu Items
    mnuOpen.Enabled = False
    mnuReports.Enabled = False
    mnuTables.Enabled = False
    mnuInput.Enabled = False
    mnuFinancial.Enabled = False
    mnuPlan.Enabled = False
    mnuProduction.Enabled = False
    mnuSecurity.Enabled = False
    cmdSecurity.Enabled = False
    cmdGroups.Enabled = False
    cmdOpen.Enabled = False
    If goUser.UserName = vbNullString Then Exit Sub
    
    Set oSecurity = New clsSecurity
    '/* Enable Secure Menu Items based on User's rights
    If oSecurity.GetMembership(goUser.UserName, vData, goApplication.SecurityDatabasePath) Then
        For i = 0 To UBound(vData)
            Select Case vData(i)
            Case "Maintenance"
                mnuOpen.Enabled = True
                mnuReports.Enabled = True
                cmdOpen.Enabled = True
            Case "Planning"
                mnuInput.Enabled = True
                mnuPlan.Enabled = True
            Case "Production"
                mnuInput.Enabled = True
                mnuProduction.Enabled = True
            Case "Reports"
                mnuReports.Enabled = True
            Case "Financial"
                mnuFinancial.Enabled = True
                mnuReports.Enabled = True
            Case "Adminstrator"
                mnuOpen.Enabled = True
                mnuReports.Enabled = True
                mnuTables.Enabled = True
                mnuInput.Enabled = True
                mnuFinancial.Enabled = True
                mnuPlan.Enabled = True
                mnuProduction.Enabled = True
                mnuSecurity.Enabled = True
                cmdSecurity.Enabled = True
                cmdGroups.Enabled = True
                cmdOpen.Enabled = True
            End Select
        Next i
    End If
    
    Erase vData
    Set oSecurity = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGroups_Click()
    cfSecurityGroups.Show
End Sub

Private Sub mnuManageSecurity_Click()
    cfSecurityManage.Show
End Sub


Private Sub mnuOpen_Click()
    frmViewDatabase.Show
End Sub


