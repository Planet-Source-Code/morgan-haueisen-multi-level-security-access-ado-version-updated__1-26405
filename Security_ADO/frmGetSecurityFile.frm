VERSION 5.00
Begin VB.Form frmGetSecurityFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Security file to modify"
   ClientHeight    =   4770
   ClientLeft      =   2355
   ClientTop       =   2160
   ClientWidth     =   7635
   Icon            =   "frmGetSecurityFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   450
      Left            =   5160
      TabIndex        =   3
      Top             =   3975
      Width           =   2010
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   5070
      Pattern         =   "*.pwd"
      TabIndex        =   2
      Top             =   495
      Width           =   2205
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   135
      Width           =   7035
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   0
      Top             =   510
      Width           =   4770
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Security file to modify"
      Height          =   375
      Left            =   255
      TabIndex        =   4
      Top             =   3900
      Width           =   4500
   End
End
Attribute VB_Name = "frmGetSecurityFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If File1.FileName = vbNullString Then
        MsgBox "No Security file selected"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Call InitSecurity(App.Path & "\WorkingDB.mdb", File1.Path & "\" & File1.FileName)
    Me.Hide
    DoEvents
    cfSecurityGroups.Show vbModal
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub Form_Load()
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGetSecurityFile = Nothing
    End
End Sub


