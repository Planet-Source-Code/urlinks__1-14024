Attribute VB_Name = "Module2"
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Dim OFName As OPENFILENAME
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Function ShowSave(Phwnd As Long, Optional FileFilter As String = "", Optional strInitialDir As String = "c:\") As String
    'Set the structure size
    Dim i As Integer
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hwndOwner = Phwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    If FileFilter = "" Then
        OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    Else
        OFName.lpstrFilter = FileFilter & "Files" & "(*." & FileFilter & ")" + Chr$(0) + "*." & FileFilter + Chr$(0) + "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    End If
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = strInitialDir
    'Set the dialog title
    OFName.lpstrTitle = "Save File"
    'no extra flags
    OFName.flags = 0

    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
        i = InStr(1, OFName.lpstrFile, Chr(0))
        If i > 0 Then
            ShowSave = Mid(OFName.lpstrFile, 1, i - 1)
        Else
            ShowSave = Trim$(OFName.lpstrFile)
        End If
    Else
        ShowSave = ""
    End If
End Function
