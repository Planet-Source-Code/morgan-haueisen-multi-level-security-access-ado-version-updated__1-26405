VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public oADODBConnection As ADODB.Connection

Public Enum e_DatabaseTypes
    dbt_Undefined = 0
    dbt_OracleMSDA = 1
    dbt_OracleODBC = 2
    dbt_SQLserver = 3
    dbt_MicrosoftAccessFile = 5
    dbt_MicrosoftAccess97File = 6
    dbt_MicrosoftAccess2KFile = 7
    dbt_DSNFile = 8
End Enum

Public Function BuildConnectString(ByVal DatabaseType As e_DatabaseTypes, ByVal ServerOrFilename As String, Optional ByVal databaseName As String, Optional ByVal UserName As String, Optional ByVal Password As String) As String
    Select Case DatabaseType
        Case dbt_OracleMSDA
            BuildConnectString = "Provider=MSDAORA;Data Source=" & ServerOrFilename & ";User ID=" & IIf(UserName <> "", UserName, "") & ";Password=" & IIf(Password <> "", Password, "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case dbt_OracleODBC
            BuildConnectString = "DRIVER={Microsoft ODBC for Oracle};SERVER=" & ServerOrFilename & ";UID=" & UserName & ";PWD=" & Password & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case dbt_SQLserver
            BuildConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & ServerOrFilename & ";User ID=" & IIf(UserName <> "", UserName, "") & ";Password=" & IIf(Password <> "", Password, "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case dbt_DSNFile
            BuildConnectString = "Provider=MSDASQL;DSN=" & ServerOrFilename & ";UID=" & IIf(UserName <> "", UserName, "") & ";PWD=" & IIf(Password <> "", Password & ";", "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case dbt_MicrosoftAccess2KFile, dbt_MicrosoftAccess97File, dbt_MicrosoftAccessFile
            BuildConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & ServerOrFilename & ";DefaultDir=" & RetPathOnly(ServerOrFilename) & ";PWD=" & IIf(Password <> "", Password & ";", ";")
    End Select
End Function

Private Function RetPathOnly(inputString As String) As String
  Dim j As Integer
    j = InStrRev(inputString, "\", , vbTextCompare)
    RetPathOnly = Mid(inputString, 1, j)
End Function

Public Function adoConnectOpen(oADODBConnection As ADODB.Connection, _
                                DB_Type As e_DatabaseTypes, _
                                ByVal ServerOrFilename As String, _
                                Optional ByVal dbPathFilename As String = "", _
                                Optional CommandType As CommandTypeEnum = adCmdStoredProc, _
                                Optional CursorLocation As CursorLocationEnum = adUseClient, _
                                Optional ByVal UserName As String = "", _
                                Optional ByVal Password As String = "") As Boolean
    
    
    On Error GoTo adoConnectError
    
    Set oADODBConnection = New ADODB.Connection

    If dbPathFilename = vbNullString Then dbPathFilename = ServerOrFilename
    
    With oADODBConnection
        .CursorLocation = CursorLocation '/* default = adUseClient(3)
        .Open BuildConnectString(DB_Type, ServerOrFilename, dbPathFilename, UserName, Password)
    End With

    adoConnectOpen = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectOpen = False
    
End Function


Public Function adoConnectClose(oADODBConnection As ADODB.Connection) As Boolean
    
    On Error GoTo adoConnectError
    oADODBConnection.Close
    adoConnectClose = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectClose = False
    
End Function



