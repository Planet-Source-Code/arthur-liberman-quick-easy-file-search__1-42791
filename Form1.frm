VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "File Search"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   525
      Left            =   1320
      TabIndex        =   9
      Top             =   30
      Width           =   1305
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Text            =   "Search In..."
      Top             =   630
      Width           =   6795
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6420
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   5070
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   8943
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Files"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dirs"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   990
      Width           =   6795
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   630
      TabIndex        =   2
      Top             =   4260
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Hidden          =   -1  'True
      Left            =   360
      System          =   -1  'True
      TabIndex        =   3
      Top             =   4260
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   450
      TabIndex        =   1
      Top             =   4170
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find files"
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1305
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2790
      TabIndex        =   7
      Top             =   2970
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuRF 
         Caption         =   "&Run File"
      End
      Begin VB.Menu mnuOF 
         Caption         =   "&Open Folder"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************************
'*This program is meant to demonstrate how quickly and easily you can make   *
'*a really fast and effective search engine for pretty much any kind of file.*
'*****************************************************************************
'*Please if plan to use this code credit me :)                               *
'*****************************************************************************
'*                        Written by: Arthur Liberman                        *
'*****************************************************************************

Dim I As Long, G As Long, PathSrc As String, Dirs() As String, StopSrch As Boolean

Private Sub Command1_Click()
If Combo1.ListIndex = -1 Then Exit Sub 'Ignore search if no drive selected
'Start the search
Dim I As Integer
StopSrch = False 'Reset "Stop" ID so it won't stop before it even started :p
Command1.Enabled = False 'disable button so there will be no simultanious searching
Command2.Enabled = True 'Enable Stop button so you can stop search
LV1.ListItems.Clear 'clear the file list if has been searched before
Form1.Refresh
If Combo1.ListIndex = Combo1.ListCount - 1 Then 'check if picked search all HDs
    For I = 0 To Drive1.ListCount - 1
        If DrvType(Drive1.List(I)) = DRIVE_FIXED Then
            'Call the search procedure for each non-removable drive
            Call DoSearch(UCase(Left(Drive1.List(I), 2)) & "\")
        End If
    Next I
Else
    Call DoSearch(UCase(Left(Combo1.List(Combo1.ListIndex), 2)) & "\") 'else, search selected drive
End If
Command1.Enabled = True 'Enable the button for next search
Command2.Enabled = False 'Disable the button as there's no more search going on
'specify the amount of found files
SB1.Panels(1).Text = LV1.ListItems.Count & " File(s) found"
End Sub

Public Sub DoSearch(SearchPath As String)
On Error Resume Next
ReDim Dirs(0)
Dir1.Path = SearchPath
If Err.Number <> 0 Then
    MsgBox "An error has occured: " & Err.Description, vbCritical, "Error No." & Err.Number
    Exit Sub
End If
If Text1.Text = "" Then 'check for the "search for"
    File1.Pattern = "*.*" 'if empty search for everything
Else
    File1.Pattern = Text1.Text 'if not search for specified file types
End If
If Len(Dir1.Path) = 3 Then
    Dirs(0) = SearchPath 'if root directory add to directory list
    Call AddFiles 'add all files in root dir
End If
For I = 0 To Dir1.ListCount - 1 'do for all directories in the root dir
    If I > 0 Then
        'redim only if this item does not exist yet
        ReDim Preserve Dirs(UBound(Dirs) + 1)
    End If
    Dirs(UBound(Dirs)) = Dir1.List(I) 'add directory to list
    Call AddFiles 'add all files in the current directory
Next I
Call AddDirs 'find all sub directories on the drive
End Sub

Public Sub AddDirs()
Dim I As Long, G As Long
I = 0
Dir1.Path = Dirs(I) 'pass the directory to search for other directories in
'do until found & recorded all directories on the drive
Do Until I = UBound(Dirs) And Dir1.ListCount = 0
    For G = 0 To Dir1.ListCount - 1 'do for each sub directory
        'resize the dynamic array to fit last object
        ReDim Preserve Dirs(UBound(Dirs) + 1)
        Dirs(UBound(Dirs)) = Dir1.List(G) 'add sub-directory to list
        SB1.Panels(1).Text = Dir1.List(G) 'show directory path and name
        If StopSrch = True Then Exit Sub 'If pressed stop, then stop
        Call AddFiles 'add all files in that directory
        DoEvents
    Next G
    I = I + 1
    Dir1.Path = Dirs(I) 'go to next directory in the list
Loop
End Sub

Public Sub AddFiles()
File1.Path = Dirs(UBound(Dirs)) 'Specify directory to search in
If Len(File1.Path) = 3 Then
    PathSrc = File1.Path 'if root dir don't add "\"
Else
    PathSrc = File1.Path & "\" 'if not root dir add "\"
End If
'add all the existing files to the list
For G = 0 To File1.ListCount - 1
    LV1.ListItems.Add LV1.ListItems.Count + 1, "", File1.List(G) 'Add file name
    LV1.ListItems(LV1.ListItems.Count).SubItems(1) = PathSrc 'Add file path
    'Add file size
    LV1.ListItems(LV1.ListItems.Count).SubItems(2) = FileSize(FileLen(PathSrc & File1.List(G)))
    'Add file time and date
    LV1.ListItems(LV1.ListItems.Count).SubItems(3) = FileDate(PathSrc & File1.List(G))
Next G
End Sub

Private Sub Command2_Click()
StopSrch = True 'Stop search
End Sub

Private Sub Form_Load()
Dim DrvStr As String, HDDCnt As Integer
DrvStr = ""
For I = 0 To Drive1.ListCount - 1
    Combo1.AddItem UCase(Drive1.List(I)) 'Add drive to the list
    'if non-removable drive, add it to non-removable drive string.
    If DrvType(Drive1.List(I)) = DRIVE_FIXED Then
        DrvStr = DrvStr & UCase(Left(Drive1.List(I), 2)) & ", "
        HDDCnt = HDDCnt + 1
    End If
Next I
'only add to list if there are more than one fixed disk
If HDDCnt > 1 Then
    DrvStr = Left(DrvStr, Len(DrvStr) - 2) 'remove the ", " part of the string
    Combo1.AddItem DrvStr 'add as last itme
End If
End Sub

Private Sub Form_Resize()
'resize the controls so the form will stil look good, even if its resized :)
If Form1.WindowState = vbMinimized Then Exit Sub
SB1.Panels(1).Width = Form1.ScaleWidth - 250
If Form1.Height < 2700 Then Form1.Height = 2700
LV1.Height = Form1.ScaleHeight - 1695
LV1.Width = Form1.ScaleWidth - 15
Text1.Width = LV1.Width
Combo1.Width = Text1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
LV1.ListItems.Clear 'unload objects, free memory
Unload Form1 'exit
End
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
LV1.SortKey = ColumnHeader.Index - 1 'Tell by which column to sort
'Reverse Sort order
If LV1.SortOrder = lvwAscending Then LV1.SortOrder = lvwDescending Else LV1.SortOrder = lvwAscending
LV1.Sorted = True 'Sort!
End Sub

Private Sub LV1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuFile 'If right button clicked, show popup menu
End Sub

Private Sub mnuOF_Click()
'Open the folder
If LV1.ListItems.Count > 0 Then Call ShellExecute(Form1.hwnd, "Open", LV1.ListItems(LV1.SelectedItem.Index).SubItems(1), "", "", 1)
End Sub

Private Sub mnuRF_Click()
'Run the file
If LV1.ListItems.Count > 0 Then Call ShellExecute(Form1.hwnd, "Open", LV1.ListItems(LV1.SelectedItem.Index).SubItems(1) & LV1.ListItems(LV1.SelectedItem.Index).Text, "", "", 1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click 'If pressed "Return" key do search
End Sub
