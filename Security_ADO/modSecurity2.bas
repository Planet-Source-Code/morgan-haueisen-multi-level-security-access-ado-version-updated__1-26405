Attribute VB_Name = "modSecurity"
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2000-2001

'/* SetUp Form:
'/*   Used to setup groups in a new application.
'/*   Not to be included in the final application.
'/*         cfSecurityGroups.frm

'/* Required Forms:
'/*         cfSecurityLogin.frm
'/*         cfSecurityManager.frm
'/* Required Class:
'/*         cSecurity.cls
'/* Required File
'/*         security.pwd

Option Explicit

Public Const mstrAccessProvider351         As String = "Provider= Microsoft.Jet.OLEDB.3.51;"
Public Const mstrAccessProvider40          As String = "Provider=Microsoft.Jet.OLEDB.4.0;"

'/* For password protected database file (if required) */
Public Const DB_PWD As String = "123456"

Public Type goUserType
    UserName As String
    Password As String
End Type
Public goUser As goUserType

Public Type goApplicationType
    SourceDatabasePath As String
    SecurityDatabasePath As String
    SystemID As String
End Type
Public goApplication As goApplicationType

Private Sub SetMenuAccess_Example()
'  Dim oSecurity As clsSecurity
'  Dim vData() As Variant
'  Dim i As Integer
'
'    '/* Disable Secure Menu Items
'    mnuFile.Enabled = False
'    mnuReports.Enabled = False
'    mnuTables.Enabled = False
'    mnuInput.Enabled = False
'    mnuFinancial.Enabled = False
'    mnuPlan.Enabled = False
'    mnuShiftInfo.Enabled = False
'    mnuSecurity.Enabled = False
'    If goUser.UserName = vbNullString Then Exit Sub
'
'    Set oSecurity = New clsSecurity
'    '/* Enable Secure Menu Items based on User's rights
'    If oSecurity.GetMembership(goUser.UserName, vData, goApplication.SecurityDatabasePath) Then
'        For i = 0 To UBound(vData)
'            Select Case vData(i)
'            Case "Maintenance"
'                mnuFile.Enabled = True
'                mnuReports.Enabled = True
'            Case "Planning"
'                mnuInput.Enabled = True
'                mnuPlan.Enabled = True
'            Case "Production"
'                mnuFile.Enabled = True
'                mnuInput.Enabled = True
'                mnuShiftInfo.Enabled = True
'            Case "Reports"
'                mnuReports.Enabled = True
'            Case "Financial"
'                mnuFinancial.Enabled = True
'                mnuReports.Enabled = True
'            Case "Adminstrator"
'                mnuFile.Enabled = True
'                mnuReports.Enabled = True
'                mnuTables.Enabled = True
'                mnuInput.Enabled = True
'                mnuFinancial.Enabled = True
'                mnuPlan.Enabled = True
'                mnuShiftInfo.Enabled = True
'                mnuSecurity.Enabled = True
'            End Select
'        Next i
'    End If
'
'    Erase vData
'    Set oSecurity = Nothing

End Sub


Public Sub InitSecurity(ByVal MDBfile As String, ByVal MDWfile As String, Optional ByVal MDWSystemID As String = "")
    MDWSystemID = "morganh" & Trim(MDWSystemID)
    goApplication.SourceDatabasePath = MDBfile
    goApplication.SecurityDatabasePath = MDWfile
    goApplication.SystemID = MDWSystemID
    
    If Dir$(MDWfile) = vbNullString Then
        MsgBox "The Security file is missing.  Please contact your system adminstrator for assistance", vbCritical
        End
    ElseIf Dir$(MDBfile) = vbNullString Then
        MsgBox "The Database file is missing.  Please contact your system adminstrator for assistance", vbCritical
        End
    End If
    
End Sub

