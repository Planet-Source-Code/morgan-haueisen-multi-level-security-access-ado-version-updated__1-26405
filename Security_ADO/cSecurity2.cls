VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2000-2001

'/* This class is the server for the client forms and modules
'/* that maintain user and group security.

'/* cfSecurityGroups.frm
'/* cfSecurityLogin.frm
'/* cfSecurityManage.frm
'/* modSecurity.bas
'/* cADOConnect.cls
'/* cADORecordset.cls

Option Explicit

'/* Password for Security file (security.pwd)
Const DB_PWD As String = "MorganH"

Public Function EditUser(ByVal UserOldName As String, ByVal UserNewName As String, ByVal NoExp As Integer, _
                            ByVal SecurityDatabasePath As String, _
                            ByVal SystemID As String) As Boolean
'/* This function creates a user
'/* UserOldName - A string that holds the name of the selected user.
'/* UserNewName - A string that holds the new name of the selected user.
'/* NoExp    - A integer that determins if a user's password expires.
'/* SystemID - A string that holds the applications ID. This prevents the
'/*            security file from being copied from one app to another to gain access
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.
    
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    If UserNewName = vbNullString Then UserNewName = UserOldName
    
    MySet.OpenIt "Select * From Users Where [UserName] = '" & UserOldName & "'", MyDB
    If Not (MySet.EOF And MySet.BOF) Then
        MySet.Edit
        MySet!UserName = UserNewName
        MySet!NoExp = CBool(NoExp)
        MySet.Update
        EditUser = True
    Else
        EditUser = False
    End If
    
    MySet.CloseIt
    
    If UserOldName <> UserNewName Then
        MyDB.Execute "UPDATE UserRights SET UserRights.UserName = '" & UserNewName & "' WHERE (((UserRights.UserName)='" & UserOldName & "'));"
    End If

    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::EditUser", Err.Description
    EditUser = False
    Resume Exit_Function
   
End Function


Public Function ListAllGroups(ObjBox As Object, ByVal SecurityDatabasePath As String, Optional ByVal GroupName As String = "") As Variant
'/* This function lists all groups defined with the current
'/* security file and places values into an ListBox or ComboBox.

'/* ObjBox - ListBox or ComboBox the info will be put in.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim SetIndex As Boolean, i As Integer
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    ObjBox.Clear
    
    MySet.OpenIt "Select * From Groups;", MyDB
    If MySet.EOF And MySet.BOF Then
    Else
        Do
            ObjBox.AddItem MySet!GroupName
            If MySet!GroupName = GroupName Then SetIndex = True
            MySet.MoveNext
        Loop Until MySet.EOF
    End If
    
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    If SetIndex Then
        For i = 0 To ObjBox.ListCount - 1
            If ObjBox.List(i) = GroupName Then
                ObjBox.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    ListAllGroups = True
    
Exit_Function:
    On Local Error Resume Next
Exit Function
    
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::ListAllGroups", Err.Description
    ListAllGroups = False
    Resume Exit_Function
    
End Function

Public Function ListAllGroupsPerUser(ObjBoxNotIn As Object, ObjBoxIn As Object, ByVal UserName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will fill a listbox of all of the groups that a User
'/* is not part of and a second listbox that contains the Groups
'/* that the User is a member of.

'/* ObjBoxNotIn - Listbox or ComoBox that contains groups the user is not a member of.
'/* ObjBoxIn - Listbox or ComoBox that contains groups the user is a member of.
'/* UserName - a string that holds the name of the user.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim MyDB As ADODB.Connection
  Dim MyGroups As New clsADORecordset
  Dim MyRights As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    ObjBoxIn.Clear
    ObjBoxNotIn.Clear
    
    MyRights.OpenIt "Select * From UserRights Where UserRights.UserName='" & UserName & "';", MyDB
    MyGroups.OpenIt "Select * From Groups;", MyDB
    
    If MyGroups.EOF And MyGroups.BOF Then
    Else
        Do
            MyRights.FindFirst "[GroupName]='" & MyGroups!GroupName & "'"
            If MyRights.NoMatch Then
                ObjBoxNotIn.AddItem MyGroups!GroupName
            Else
                ObjBoxIn.AddItem MyGroups!GroupName
            End If
            MyGroups.MoveNext
        Loop Until MyGroups.EOF
    End If
    
    MyGroups.CloseIt
    MyRights.CloseIt
    MyDB.Close
    Set MyGroups = Nothing
    Set MyRights = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
End Function

Public Function CreateUser(ByVal UserName As String, ByVal NoExp As Integer, _
                            ByVal SecurityDatabasePath As String, _
                            ByVal SystemID As String) As Boolean
'/* This function creates a user
'/* UserName - A string that holds the name of the user.
'/* Password - A string that holds the user's password.
'/* SystemID - A string that holds the applications ID. This prevents the
'/*            security file from being copied from one app to another to gain access
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.
    
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From Users Where [UserName] = '" & UserName & "'", MyDB
    If MySet.EOF And MySet.BOF Then
        MySet.AddNew
        MySet!UserName = UserName
        MySet!SystemID = SystemID
        MySet!Date = Date
        MySet!NoExp = CBool(NoExp)
        MySet!UserPassword = "password"
        MsgBox "The default password has been set to the word 'password'.", vbInformation
        MySet.Update
        CreateUser = True
    Else
        CreateUser = False
    End If
    
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::CreateUser", Err.Description
    CreateUser = False
    Resume Exit_Function
   
End Function

Public Function CreateGroup(ByVal NewGroupName As String, ByVal SecurityDatabasePath As String, Optional ByVal GroupName As String = "") As Boolean
'/* This function creates a user
'/* GroupName - A string that holds the name of the user.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.
    
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    If GroupName > vbNullString Then '/* Edit Group Name
        MySet.OpenIt "Select Groups.* From Groups Where [GroupName] = '" & NewGroupName & "'", MyDB
        If MySet.EOF And MySet.BOF Then
            MySet.CloseIt
            
            MySet.OpenIt "Select Groups.* From Groups Where [GroupName] = '" & GroupName & "'", MyDB
            If MySet.EOF And MySet.BOF Then
                '/* do nothing
            Else
                MySet.Edit
                MySet!GroupName = NewGroupName
                MySet.Update
                CreateGroup = True
            End If
            
            MySet.CloseIt
            MyDB.Execute "UPDATE UserRights SET UserRights.GroupName ='" & NewGroupName & _
                            "' WHERE UserRights.GroupName='" & GroupName & "';"
            
        Else
            MySet.CloseIt
            CreateGroup = False
        End If
    Else '/* Add a Group Name
        MySet.OpenIt "Select Groups.* From Groups Where [GroupName] = '" & NewGroupName & "'", MyDB
        If MySet.EOF And MySet.BOF Then
            MySet.AddNew
            MySet!GroupName = NewGroupName
            MySet.Update
            CreateGroup = True
        Else
            CreateGroup = False
        End If
        MySet.CloseIt
    End If
    
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::CreateUser", Err.Description
    CreateGroup = False
Resume Exit_Function
   
End Function


Public Function AddUserToGroup(ByVal GroupName As String, ByVal UserName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function adds a user to a group.
'/* UserName - A string that holds the name of the user.
'/* GroupName - A string that holds the group name.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  'Dim MyDB  As Database
  'Dim MySet As Recordset
  
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect

    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From UserRights Where [GroupName] = '" & GroupName & "'", MyDB
    
    MySet.FindFirst "[UserName] = '" & UserName & "'"
    If MySet.NoMatch Then
        MySet.AddNew
        MySet!UserName = UserName
        MySet!GroupName = GroupName
        MySet.Update
        AddUserToGroup = True
    End If
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
Exit_Function:
    On Error Resume Next
Exit Function
    
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::AddUserToGroup", Err.Description
    Resume Exit_Function
End Function

Public Function Login(ByVal UserName As String, ByVal Password As String, _
                        ByVal SecurityDatabasePath As String, _
                        ByVal SystemID As String) As Integer
'/* This function will login to the database and the security file.
'/* UserName - A string that holds the name of the user.
'/* Password - A string that holds the user's password.
'/* SystemID - A string that holds the applications ID. This prevents the
'/*            security file from being copied from one app to another to gain access.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True or False
    
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
  Dim ExpDays As Byte
  Dim PassDiff As Integer
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select ExpDays.Days from ExpDays", MyDB
    ExpDays = MySet!Days
    MySet.CloseIt
    
    MySet.OpenIt "Select * From Users Where [UserName] = '" & UserName & _
                                    "' AND [SystemID]='" & SystemID & "';", MyDB
    
    If MySet.EOF And MySet.BOF Then
        '/* No user ID found
        Login = 0
    Else
        '/* if the user's password is password then ignore any password the user has typed.
        If MySet!UserPassword = "password" Then Password = "password"
        If MySet!UserPassword <> Password Then
            Login = 0
        Else
            '/* The login was successful.
            PassDiff = DateDiff("d", MySet!Date, Date) + 1
            If ExpDays = 0 Then
                Login = -1
            Else
                If Password = "password" Then
                    Login = -2
                ElseIf Not MySet!NoExp And PassDiff > ExpDays Then
                    Login = PassDiff
                Else
                    Login = -1
                End If
            End If
        End If
    End If
    
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::Login", Err.Description
    Resume Exit_Function

End Function

Public Function ExpDaysGet(ByVal SecurityDatabasePath As String) As Integer
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select ExpDays.Days from ExpDays", MyDB
    ExpDaysGet = CStr(MySet!Days)
    MySet.CloseIt
    MyDB.Close
    
    Set MyDB = Nothing
    Set MySet = Nothing

End Function

Public Sub ExpDaysSet(ByVal ExpDays As Variant, ByVal SecurityDatabasePath As String)
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, SecurityDatabasePath, , , , , DB_PWD
    MySet.OpenIt "Select ExpDays.Days from ExpDays", MyDB
    
    MySet.Edit
    MySet!Days = ExpDays
    MySet.Update
    
    MySet.CloseIt
    MyDB.Close
    
    Set MyDB = Nothing
    Set MySet = Nothing

End Sub

Public Function RemoveUserFromGroup(ByVal GroupName As String, ByVal UserName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function removes a user from a group.
'/* UserName - A string that holds the name of the user.
'/* GroupName - A string that holds the group name.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim MyDB As ADODB.Connection
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MyDB.Execute "DELETE UserRights.* From UserRights WHERE " & _
            "UserRights.UserName='" & UserName & "' AND UserRights.GroupName='" & GroupName & "';"
    
    MyDB.Close
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    RemoveUserFromGroup = True
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::RemoveUserFromGroup", Err.Description
    Resume Exit_Function

End Function

Public Function ListAllUsers(ObjBox As ComboBox, ByVal SecurityDatabasePath As String, Optional ByVal UserName As String = "") As Boolean
'/* This function list all users defined

'/* ObjBox - ListBox or ComboBox the info will be put in.
'/* UserName - Set the listindex to a user's name.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim i     As Integer
  Dim j     As Integer
  Dim SetIndex As Boolean
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    ObjBox.Clear
    
    MySet.OpenIt "Select Users.* From Users;", MyDB
    If MySet.EOF And MySet.BOF Then
    Else
        Do
            ObjBox.AddItem MySet!UserName
            ObjBox.ItemData(ObjBox.NewIndex) = CInt(MySet!NoExp)
            If MySet!UserName = UserName Then SetIndex = True
            MySet.MoveNext
        Loop Until MySet.EOF
    End If
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    If SetIndex Then
        For i = 0 To ObjBox.ListCount - 1
            If ObjBox.List(i) = UserName Then
                ObjBox.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    ListAllUsers = True

Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::ListAllUsers", Err.Description
    Resume Exit_Function

End Function

Public Function ChangePassword(ByVal UserName As String, _
                                ByVal NewPassword As String, _
                                ByVal NewPasswordVerify As String, _
                                ByVal Password As String, _
                                ByVal SecurityDatabasePath As String, _
                                Optional DisplayErrors As Boolean = False)

'/* This function will change the password for a user account.
'/* UserName - This string contains the name of the user whose name will be changed.
'/* NewPassword - This string contains the new password that the user entered.
'/* Pasword - This string contains the user's old password.
'/* NewPasswordVerify -This string contains the value the user entered
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* DisplayErrors - Display msgbox errors
'/*   to verify the new password.

  'Dim MyDB  As Database
  'Dim MySet As Recordset
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect

    
    On Error GoTo Error_Function
    
    If NewPassword <> NewPasswordVerify Then
        If DisplayErrors Then MsgBox "New password does not match the verify password.", vbCritical
        Exit Function
    End If
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From Users Where [UserName] = '" & UserName & "'", MyDB
    If MySet.EOF And MySet.BOF Then
        If DisplayErrors Then MsgBox "The user name or password is invalid.", vbCritical
    Else
        If MySet!UserPassword = Password Then
            MySet.Edit
            MySet!UserPassword = NewPassword
            MySet!Date = Date
            MySet.Update
            ChangePassword = True
        ElseIf Password = "ChangePassword" Then
            MySet.Edit
            MySet!UserPassword = NewPassword
            MySet!Date = Date
            MySet.Update
            ChangePassword = True
        Else
            If DisplayErrors Then MsgBox "The user name or password is invalid.", vbCritical
        End If
    End If

    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing

    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::ChangePassword", Err.Description
    ChangePassword = False
Resume Exit_Function
    
End Function

Public Function DeleteUser(ByVal UserName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will delete a user from the current security file.
'/* UserName - This string contains the name of the user.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MyDB.Execute "DELETE Users.* From Users WHERE Users.UserName='" & UserName & "';"
    MyDB.Execute "DELETE UserRights.* From UserRights WHERE UserRights.UserName='" & UserName & "';"
    
    MyDB.Close
    Set MyDB = Nothing
    
    DeleteUser = True
    
Exit_Function:
    On Error Resume Next
Exit Function
    

Error_Function:
    Err.Raise Err.Number, "clsSecurity::DeleteUser", Err.Description
    Resume Exit_Function
End Function

Public Function ClearPassword(ByVal UserName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will clear the password for a user account
'/* to the word password.
'/* UserName - This string contains the name of the user whose name will be changed.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.
    
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From Users Where [UserName] = '" & UserName & "'", MyDB
    
    If MySet.EOF And MySet.BOF Then
    Else
        MySet.Edit
        MySet!UserPassword = "password"
        MySet!Date = Date
        MySet.Update
        ClearPassword = True
        'MsgBox "The password has been reset to the word password.", vbInformation
    End If
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::ClearPassword", Err.Description
    ClearPassword = False
Resume Exit_Function

End Function


Public Function IsMember(ByVal UserName As String, ByVal GroupName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will return true if the user is in the specified group.
'/* UserName - a string that holds the name of the User.
'/* GroupName - a string that holds the name of the Group.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From UserRights Where [UserName] = '" & UserName & "' AND [GroupName] = '" & GroupName & "'", MyDB
    If MySet.EOF And MySet.BOF Then
        IsMember = False
    Else
        IsMember = True
    End If
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::IsMember", Err.Description
    Resume Exit_Function

End Function
Public Function GetMembership(ByVal UserName As String, vData As Variant, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will return an array of user rights.
'/* UserName - a string that holds the name of the user.
'/* vData - a variant array of the user's rights
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim Count As Integer
  Dim i As Integer
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select * From UserRights Where [UserName] = '" & UserName & "';", MyDB
    If MySet.EOF And MySet.BOF Then
    Else
        MySet.MoveLast
        Count = MySet.RecordCount - 1
        ReDim vData(Count)
        MySet.MoveFirst
        Do
            vData(i) = MySet!GroupName
            i = i + 1
            MySet.MoveNext
        Loop Until MySet.EOF
        GetMembership = True
    End If
    MySet.CloseIt
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing
    Set OpenDB = Nothing
    
    
Exit_Function:
    On Error Resume Next
Exit Function
    
Error_Function:
    Err.Raise Err.Number, "clsSecurity::IsMember", Err.Description
    Resume Exit_Function

End Function


Public Function DeleteGroup(ByVal GroupName As String, ByVal SecurityDatabasePath As String) As Boolean
'/* This function will delete a user from the current security file.
'/* GroupName - This string contains the name of the group.
'/* SecurityDatabasePath - The path and file name of the security database file.
'/* Returns:    True/False if ok.

  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    On Error GoTo Error_Function
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, goApplication.SecurityDatabasePath, , , , , DB_PWD
    
    MyDB.Execute "DELETE Groups.* From Groups WHERE Groups.GroupName='" & GroupName & "';"
    MyDB.Execute "DELETE UserRights.* From UserRights WHERE UserRights.GroupName='" & GroupName & "';"

    MyDB.Close
    Set MyDB = Nothing
    
    DeleteGroup = True
    
Exit_Function:
    On Error Resume Next
Exit Function
    

Error_Function:
    Err.Raise Err.Number, "clsSecurity::DeleteUser", Err.Description
    Resume Exit_Function

End Function

Public Function SystemIDGet(ByVal SecurityDatabasePath As String) As String
  Dim MyDB As ADODB.Connection
  Dim MySet As New clsADORecordset
  Dim OpenDB As New clsADOConnect
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, SecurityDatabasePath, , , , , DB_PWD
    
    MySet.OpenIt "Select Users.SystemID from Users", MyDB
    SystemIDGet = Mid(MySet!SystemID, 8) & vbNullString
    MySet.CloseIt
    MyDB.Close
    
    Set MyDB = Nothing
    Set MySet = Nothing

End Function
Public Sub SystemIDSet(ByVal SystemID As String, ByVal SecurityDatabasePath As String)
  Dim MyDB As ADODB.Connection
  Dim OpenDB As New clsADOConnect
    
    OpenDB.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, SecurityDatabasePath, , , , , DB_PWD
    
    SystemID = "morganh" & Trim(SystemID)
    MyDB.Execute "UPDATE Users SET Users.SystemID = '" & SystemID & "';"

    MyDB.Close
    
    Set MyDB = Nothing

End Sub

