VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "URLinks"
   ClientHeight    =   7236
   ClientLeft      =   132
   ClientTop       =   360
   ClientWidth     =   8892
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   8892
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1716
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   5040
      Width           =   8544
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   228
      TabIndex        =   4
      Top             =   816
      Width           =   8568
   End
   Begin VB.ListBox List1 
      Height          =   3888
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Width           =   8556
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   228
      TabIndex        =   1
      Top             =   492
      Width           =   2964
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   396
      Left            =   3816
      TabIndex        =   0
      Top             =   6840
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   192
      Left            =   2340
      TabIndex        =   3
      Top             =   108
      Width           =   36
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save As"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LocalFileName As String
Dim TextChange As Boolean
Private Sub Form_Load()

With Combo1
   .AddItem "Normal Entry"
   .ItemData(.NewIndex) = &H1
   .AddItem "Edited Entry (IE5)"
   .ItemData(.NewIndex) = &H8
   .AddItem "Offline Entry"
   .ItemData(.NewIndex) = &H10
   .AddItem "Online Entry"
   .ItemData(.NewIndex) = &H20
   .AddItem "Stick Entry"
   .ItemData(.NewIndex) = &H40
   .AddItem "Sparse Entry (n/a)"
   .ItemData(.NewIndex) = &H10000
   .AddItem "Cookies"
   .ItemData(.NewIndex) = &H100000
   .AddItem "Visited History"
   .ItemData(.NewIndex) = &H200000
   .AddItem "Default Filter"
   .ItemData(.NewIndex) = URLCACHE_FIND_DEFAULT_FILTER
   .ListIndex = 0
End With
  
End Sub


Private Sub Command1_Click()

Dim numEntries As Long
Dim cacheType As Long

cacheType = Combo1.ItemData(Combo1.ListIndex)

Label1.Caption = "Working ..."
Label1.Refresh
Text1 = ""
Text2 = ""
List1.Clear
List1.Visible = False

numEntries = GetCacheURLList(cacheType)

List1.Visible = True

Label1.Caption = Format$(List1.ListCount, sFileCount)
Label1.Caption = Format$(numEntries, sFileCount)
  
  
End Sub


Private Sub Form_Resize()
Text1.Width = Width - 500
Text2.Width = Width - 500
List1.Width = Width - 500
End Sub

Private Sub List1_Click()
If Not TextChange Then
    If List1.Text <> "" Then
        Text1.Text = List1.List(List1.ListIndex)
        LocalFileName = GetCacheEntryInfo(List1.Text)
    End If
End If
TextChange = False
End Sub


Private Function GetCacheURLList(cacheType As Long) As Long
   
  Dim ICEI As INTERNET_CACHE_ENTRY_INFO
  Dim hFile As Long
  Dim cachefile As String
  Dim nCount As Long
  Dim dwBuffer As Long
  Dim pntrICE As Long
  
 'Like other APIs, calling FindFirstUrlCacheEntry or
 'FindNextUrlCacheEntry with an insufficient buffer will
 'cause the API to fail, and its dwBuffer points to the
 'correct size required for a successful call.
  dwBuffer = 0
  
 'Call to determine the required buffer size
  hFile = FindFirstUrlCacheEntry(vbNullString, ByVal 0, dwBuffer)
  
 'both conditions should be met by the first call
  If (hFile = ERROR_CACHE_FIND_FAIL) And _
     (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
  
    'The INTERNET_CACHE_ENTRY_INFO data type is a
    'variable-length type. It is necessary to allocate
    'memory for the result of the call and pass the
    'pointer to this memory location to the API.
     pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
       
    'allocation successful
     If pntrICE Then
        
       'set a Long pointer to the memory location
        CopyMemory ByVal pntrICE, dwBuffer, 4
        
       'and call the first find API again passing the
       'pointer to the allocated memory
        hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
      
       'hfile should = 1 (success)
        If hFile <> ERROR_CACHE_FIND_FAIL Then
        
          'now just loop through the cache
           Do
           
             'the pointer has been filled, so move the
             'data back into a ICEI structure
              CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
           
             'CacheEntryType is a long representing the type of
             'entry returned, and should match our passed param.
              If (ICEI.CacheEntryType And cacheType) Then
              
                 'extract the string from the memory location
                 'pointed to by the lpszSourceUrlName member
                 'and add to a list
                  cachefile = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                  List1.AddItem cachefile
                  nCount = nCount + 1
              
              End If
              
             'free the pointer and memory associated
             'with the last-retrieved file
              Call LocalFree(pntrICE)
              
             'and again repeat the procedure, this time calling
             'FindNextUrlCacheEntry with a buffer size set to 0.
             'This will cause the call to once again fail,
             'returning the required size as dwBuffer
              dwBuffer = 0
              Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
              
             'allocate and assign the memory to the pointer
              pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
              CopyMemory ByVal pntrICE, dwBuffer, 4
              
          'and call again with the valid parameters.
          'If the call fails (no more data), the loop exits.
          'If the call is successful, the Do portion of the
          'loop is executed again, extracting the data from
          'the returned type
           Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
 
        End If 'hFile
        
     End If 'pntrICE
  
  End If 'hFile
  

 'clean up by closing the find handle, as
 'well as calling LocalFree again to be safe
  Call LocalFree(pntrICE)
  Call FindCloseUrlCache(hFile)
  
  GetCacheURLList = nCount
  
End Function

Public Function GetCacheEntryInfo(lpszUrl As String) As String
Dim dwEntrySize As Long
Dim ICEI As INTERNET_CACHE_ENTRY_INFO

Dim strTemp As String
Dim dwTemp As Long
Dim rtn As Long
Dim pntrICE As Long
dwEntrySize = 0

rtn = GetUrlCacheEntryInfo(lpszUrl, ByVal 0, dwEntrySize)
If (rtn <> ERROR_FILE_NOT_FOUND) Then
    If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
        pntrICE = LocalAlloc(LMEM_FIXED, dwEntrySize)
        If pntrICE Then
        
            CopyMemory ByVal pntrICE, dwEntrySize, 4
        
  
            rtn = GetUrlCacheEntryInfo(lpszUrl, ByVal pntrICE, dwEntrySize)
            If rtn = 1 Then
                Dim LocTime As SYSTEMTIME
                CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
                If (ICEI.dwHeaderInfoSize <> 0) Then
                    Text2 = ""
                    Text2 = Text2 & "Local File Name     : " & GetStrFromPtrA(ICEI.lpszLocalFileName) & vbCrLf
                    FileTimeToSystemTime ICEI.LastSyncTime, LocTime
                    Text2 = Text2 & "LastSyncTime        : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.LastModifiedTime, LocTime
                    Text2 = Text2 & "LastModifiedTime : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.LastAccessTime, LocTime
                    Text2 = Text2 & "LastAccessTime   : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.ExpireTime, LocTime
                    Text2 = Text2 & "ExpireTime             : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    Text2 = Text2 & "HitRate                    : " & (ICEI.dwHitRate) & vbCrLf
                    Text2 = Text2 & "FileSize                   : " & GetFileSize((ICEI.dwSizeHigh * MAX_DWORD) + ICEI.dwSizeLow) & vbCrLf
                    GetCacheEntryInfo = GetStrFromPtrA(ICEI.lpszLocalFileName)
                Else
                    Text2 = ""
                    Text2 = Text2 & "Local File Name     : " & GetStrFromPtrA(ICEI.lpszLocalFileName) & vbCrLf
                    FileTimeToSystemTime ICEI.LastSyncTime, LocTime
                    Text2 = Text2 & "LastSyncTime        : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.LastModifiedTime, LocTime
                    Text2 = Text2 & "LastModifiedTime : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.LastAccessTime, LocTime
                    Text2 = Text2 & "LastAccessTime   : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    FileTimeToSystemTime ICEI.ExpireTime, LocTime
                    Text2 = Text2 & "ExpireTime             : " & DayOfDate(LocTime.wDayOfWeek) & " ," & LocTime.wDay & "/" & LocTime.wMonth & "/" & LocTime.wYear & " " & LocTime.wHour & ":" & LocTime.wMinute & ":" & LocTime.wSecond & ":" & LocTime.wMilliseconds & vbCrLf
                    Text2 = Text2 & "HitRate                    : " & (ICEI.dwHitRate) & vbCrLf
                    Text2 = Text2 & "FileSize                   : " & GetFileSize((ICEI.dwSizeHigh * MAX_DWORD) + ICEI.dwSizeLow) & vbCrLf
                    GetCacheEntryInfo = GetStrFromPtrA(ICEI.lpszLocalFileName)
                End If
            Else
                Text2 = "No information about     : " & lpszUrl
            End If
            Call LocalFree(pntrICE)
        End If
    Else
        Text2 = "No information about     : " & lpszUrl
    End If
End If
End Function
        
Public Function DeleteCacheEntryInfo(lpszUrl As String) As String
Dim rtn As Long
rtn = DeleteUrlCacheEntry(lpszUrl)
If rtn <> 1 Then
    If Err.LastDllError = ERROR_ACCESS_DENIED Then
        DeleteCacheEntryInfo = "Access is denied"
    Else
        If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
            DeleteCacheEntryInfo = "File not found"
        Else
            DeleteCacheEntryInfo = "Unknown Error"
        End If
    End If
Else
    DeleteCacheEntryInfo = lpszUrl & " Successfuly deleted"
End If
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

  GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
  
End Function


Private Sub List1_DblClick()
MnuOpen_Click
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    If List1.Text <> "" Then
        If LocalFileName <> "" Then
            MnuSaveAs.Visible = True
            PopupMenu MnuFile, , x, y
        Else
            MnuSaveAs.Visible = False
            PopupMenu MnuFile, , x, y
        End If
    End If
End If
End Sub

Private Sub MnuDelete_Click()
Dim rtn As Long
Dim DelFile As String
If MsgBox("Do you want to remove this [ " & List1.Text & " ]  cache info ", vbYesNo, "Delete") = vbYes Then
    If DeleteCacheEntryInfo(List1.Text) = List1.Text & " Successfuly deleted" Then
        If MsgBox(List1.Text & " Successfuly deleted" & vbCrLf & "Do you want to remove the local file [ " & LocalFileName & " ] also", vbYesNo, "Delete") = vbYes Then
        List1.ListIndex = List1.ListIndex + 1
        DelFile = LocalFileName
        rtn = DeleteFile(DelFile)
        If rtn <> 0 Then
            MsgBox LocalFileName & " Deleted"
        Else
            MsgBox "Unable to  Delete " & DelFile
        End If
        
        End If
    Else
        MsgBox List1.Text & " can't deleted"
    End If
End If
End Sub

Private Sub MnuOpen_Click()
Dim rtn As Long
rtn = ShellExecute(Me.hWnd, vbNullString, List1.Text, vbNullString, "C:\", SW_SHOWNORMAL)
If rtn > 32 Then
Else
    ShellExecute Me.hWnd, vbNullString, LocalFileName, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Sub
Public Function DayOfDate(WkDay As Integer) As String
Select Case WkDay
    Case 0
        DayOfDate = "Sunday"
    Case 1
        DayOfDate = "Monday"
    Case 2
        DayOfDate = "Tuesday"
    Case 3
        DayOfDate = "Wednesday"
    Case 4
        DayOfDate = "Thursday"
    Case 5
        DayOfDate = "Friday"
    Case 6
        DayOfDate = "Saturday"
End Select
    
End Function
Public Function GetFileSize(nSize As Long) As String
Dim KB As Long, MB As Long
KB = 1024
MB = 1048576
If nSize < KB Then
    GetFileSize = nSize & " bytes"
Else
    If nSize < MB Then
        GetFileSize = Round((nSize / KB), 2) & " KB"
    Else
        GetFileSize = Round((nSize / MB), 2) & " MB"
    End If
End If


End Function

Private Sub MnuSaveAs_Click()
Dim SaveFileName As String, Folder As String, Extension As String
Dim i As Integer, RevName As String
RevName = StrReverse(LocalFileName)
i = InStr(1, RevName, "\")
Extension = Right(LocalFileName, 3)
i = Len(LocalFileName) - i
Folder = Mid(LocalFileName, 1, i)
Dim rtn As Long
SaveFileName = ShowSave(Me.hWnd, Extension, Folder)
If SaveFileName <> "" And LocalFileName <> "" Then
    RevName = StrReverse(SaveFileName)
    i = InStr(1, RevName, ".")
    If i = 0 Then
        SaveFileName = SaveFileName & "." & Extension
    End If
    rtn = CopyFile(LocalFileName, SaveFileName, 1)
    If rtn <> 0 Then
        MsgBox "File Saved as " & SaveFileName
    Else
        If MsgBox(SaveFileName & "  aleady exist" & vbCrLf & "Do you want to overwrite it ?", vbYesNo + vbExclamation, "Save file as") = vbYes Then
        rtn = CopyFile(LocalFileName, SaveFileName, 0)
        If rtn <> 0 Then
            MsgBox "File Saved as " & SaveFileName
        End If
        End If
    End If
End If
End Sub

Private Sub Text1_Change()
TextChange = True
List1.ListIndex = SendMessage(List1.hWnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
Else
    If KeyCode = 38 Then
        If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex - 1
    End If
    
End If

End Sub

