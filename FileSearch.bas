Attribute VB_Name = "FileSearch"
' Source: https://stackoverflow.com/questions/30511217/optimize-speed-of-recursive-file-search-in-subdirectories

Option Explicit

Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal lpFindFileData As LongPtr) As LongPtr
Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByVal lpFindFileData As LongPtr) As LongPtr

Private Type FILETIME
  dwLowDateTime  As Long
  dwHighDateTime As Long
End Type

Const MAX_PATH  As Long = 260
Const ALTERNATE As Long = 14

' Can be used with either W or A functions
' Pass VarPtr(wfd) to W or simply wfd to A
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime   As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime  As FILETIME
  nFileSizeHigh    As Long
  nFileSizeLow     As Long
  dwReserved0      As Long
  dwReserved1      As Long
  cFileName        As String * MAX_PATH
  cAlternate       As String * ALTERNATE
End Type

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = 16 '0x10
Private Const INVALID_HANDLE_VALUE As LongPtr = -1

Function Recurse(folderPath As String, fileName As String)
    Dim fileHandle    As LongPtr
    Dim searchPattern As String
    Dim foundPath     As String
    Dim foundItem     As String
    Dim fileData      As WIN32_FIND_DATA

    searchPattern = folderPath & "\*"

    foundPath = vbNullString
    fileHandle = FindFirstFileW(StrPtr(searchPattern), VarPtr(fileData))
    If fileHandle <> INVALID_HANDLE_VALUE Then
        Do
            foundItem = Left$(fileData.cFileName, InStr(fileData.cFileName, vbNullChar) - 1)

            If foundItem = "." Or foundItem = ".." Then 'Skip metadirectories
            'Found Directory
            ElseIf fileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                foundPath = Recurse(folderPath & "\" & foundItem, fileName)
            'Found File
            ElseIf InStr(1, foundItem, fileName, vbTextCompare) > 0 Then 'for performance
                foundPath = folderPath & "\" & foundItem
            End If

            If foundPath <> vbNullString Then
                Recurse = foundPath
                Exit Function
            End If

        Loop While FindNextFileW(fileHandle, VarPtr(fileData))
    End If

    'No Match Found
    Recurse = vbNullString
End Function

Sub TestFileSearch()

    Dim path As Variant ' цикл For Each требует этот тип
    Dim paths
    paths = Array("d:\Синица\КП\", "d:\Мякотин\САМОЕ ВАЖНОЕ\", "d:\Иншаков\КП\")
    Dim targetName As String
    Dim targetPath As String
    targetName = "КП Кореневская Солодовня 05.11.2018.xls"
'    targetPath = "D:"

    Dim target As String
    
    For Each path In paths
        target = Recurse(CStr(path), targetName)
        If target <> "" Then
            MsgBox "found: " & target
            Exit For
        End If
    Next path

    If target = "" Then MsgBox "nothing found  :( "
End Sub
