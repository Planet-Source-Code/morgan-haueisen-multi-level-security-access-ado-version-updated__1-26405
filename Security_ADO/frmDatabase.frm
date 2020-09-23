VERSION 5.00
Begin VB.Form frmViewDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmDatabase"
   ClientHeight    =   3150
   ClientLeft      =   4380
   ClientTop       =   3210
   ClientWidth     =   3960
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLast 
      Height          =   540
      Left            =   2835
      Picture         =   "frmDatabase.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1935
      Width           =   420
   End
   Begin VB.CommandButton cmdNext 
      Height          =   540
      Left            =   2385
      Picture         =   "frmDatabase.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1935
      Width           =   420
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   540
      Left            =   885
      Picture         =   "frmDatabase.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1935
      Width           =   420
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   540
      Left            =   435
      Picture         =   "frmDatabase.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1935
      Width           =   420
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "Z2U"
      DataSource      =   "Data1"
      Height          =   330
      Left            =   1965
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1155
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Z2L"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   1305
      TabIndex        =   7
      Top             =   1935
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Viewing a password protected database (WorkingDB.mdb)"
      Height          =   795
      Left            =   300
      TabIndex        =   2
      Top             =   150
      Width           =   3270
   End
End
Attribute VB_Name = "frmViewDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyDB As ADODB.Connection
Dim MySet As New clsADORecordset
Dim OpenDB As New clsADOConnect

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub


Private Sub Data1_Reposition()
    Data1.Caption = "WorkingDB.mdb (" & Data1.Recordset.AbsolutePosition + 1 & " of " & Data1.Recordset.RecordCount & ")"
End Sub

Private Sub cmdFirst_Click()
    MySet.MoveFirst
    Text1 = MySet!Z2L
    Text2 = MySet!Z2U
    lblInfo = MySet.AbsolutePosition & " of " & MySet.RecordCount
End Sub

Private Sub cmdLast_Click()
    MySet.MoveLast
    Text1 = MySet!Z2L
    Text2 = MySet!Z2U
    
    lblInfo = MySet.AbsolutePosition & " of " & MySet.RecordCount
End Sub


Private Sub cmdNext_Click()
    MySet.MoveNext
    If MySet.EOF Then MySet.MoveLast
    Text1 = MySet!Z2L
    Text2 = MySet!Z2U
    
    lblInfo = MySet.AbsolutePosition & " of " & MySet.RecordCount
End Sub

Private Sub cmdPrev_Click()
    MySet.MovePrevious
    If MySet.BOF Then MySet.MoveFirst
    Text1 = MySet!Z2L
    Text2 = MySet!Z2U

    lblInfo = MySet.AbsolutePosition & " of " & MySet.RecordCount
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SourceDatabasePath, goApplication.SourceDatabasePath, , , , DB_PWD
    MySet.OpenIt "Select * from Zones", MyDB
    
    MySet.MoveLast
    MySet.MoveFirst
    Text1 = MySet!Z2L
    Text2 = MySet!Z2U
    lblInfo = MySet.AbsolutePosition & " of " & MySet.RecordCount
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    MySet.CloseIt
    'MyDB.Close
    OpenDB.adoConnectClose MyDB
    
    Set MyDB = Nothing
    Set MySet = Nothing
    Set OpenDB = Nothing
    Set frmViewDatabase = Nothing

End Sub


