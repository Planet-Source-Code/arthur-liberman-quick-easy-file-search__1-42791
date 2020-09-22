Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_CDROM = 5

Public Const DRIVE_FIXED = 3

Public Const DRIVE_RAMDISK = 6

Public Const DRIVE_REMOTE = 4

Public Const DRIVE_REMOVABLE = 2

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Created As SYSTEMTIME, Modified As SYSTEMTIME, Accessed As SYSTEMTIME, _
W32FD As WIN32_FIND_DATA
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const MAX_PATH = 260
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Dim hFile As Long

Public Function DrvType(Drv As String) As Long
DrvType = GetDriveType(Drv) 'Return drive type
End Function

Public Function FileSize(OriginalSize As Double) As String
Dim b As Double
'If file smaller than 0.98KB
If OriginalSize < 1004 Then
    b = OriginalSize
    FileSize = b & " Bytes"
'If file larger than 0.98KB but smaller than 0.98MB
ElseIf OriginalSize >= 1004 And OriginalSize < 1027604 Then
    b = OriginalSize / 1024
    If Format(b, "###.##") = Format(b, "###.") Then
        FileSize = Format(b, "###") & " KBytes"
    Else
        FileSize = Format(b, "###.##") & " KBytes"
    End If
'If file larger than 0.98MB but smaller than 0.98GB
ElseIf OriginalSize >= 1027604 And OriginalSize < 1052266496 Then
    b = OriginalSize / 1048576
    If Format(b, "###.##") = Format(b, "###.") Then
        FileSize = Format(b, "###") & " MBytes"
    Else
        FileSize = Format(b, "###.##") & " MBytes"
    End If
'If file larger than 0.98GB
ElseIf OriginalSize >= 1052266496 Then
    b = OriginalSize / 1073741824
    If Format(b, "###.##") = Format(b, "###.") Then
        FileSize = Format(b, "###") & " GBytes"
    Else
        FileSize = Format(b, "###.##") & " GBytes"
    End If
End If
End Function

Public Function FileDate(FilePath As String) As String
hFile = 0 'Reset handle to file
hFile = FindFirstFile(FilePath, W32FD) 'Get file info and handle
Call FindClose(hFile) 'Close Find
Call FileTimeToSystemTime(W32FD.ftLastWriteTime, Modified) 'Convert FileTime into SystemTime
With Modified
    'Return the date and time
    FileDate = AddO(.wMonth) & "/" & AddO(.wDay) & "/" & .wYear & " " & AddO(.wHour) & ":" & AddO(.wMinute)
End With
End Function

Public Function AddO(ByVal InputStr As String) As String
'Supplement with the missing 0 at the start of a date/time to make it look good
If Len(InputStr) = 1 Then AddO = "0" & InputStr Else AddO = InputStr
End Function
