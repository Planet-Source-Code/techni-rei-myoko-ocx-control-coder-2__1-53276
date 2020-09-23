Attribute VB_Name = "OpenSaveDlg"
Option Explicit
Public Const imageextentions As String = "*.bmp;*.gif;*.jpg;*.jpeg;*.jpe;*.jfif;*.png"
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
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
Public SaveFileDialog As OPENFILENAME
Public OpenFileDialog As OPENFILENAME
Private rv As Long
Private sv As Long
Public Function Open_File(hwnd As Long) As String
   rv& = GetOpenFileName(OpenFileDialog)
   If (rv&) Then
      Open_File = Replace(Trim$(OpenFileDialog.lpstrFile), Chr(0), Empty)
   Else
      Open_File = ""
   End If
End Function
Public Function Save_File(hwnd As Long, defaultextention As String) As String
   sv& = GetSaveFileName(SaveFileDialog)
   Dim temp As String
   temp = ""
   If (sv&) Then
      temp = Trim$(SaveFileDialog.lpstrFile)
      temp = Left(temp, Len(temp) - 1)
      If InStrRev(temp, ".") = 0 Then temp = temp & "." & defaultextention
      If Dir(temp) <> Empty Then If MsgBox("File already exists. Do you wish to over write it?" & vbNewLine & temp, vbYesNo, "File exists") = vbNo Then temp = ""
      Save_File = temp
   End If
End Function

Public Sub InitSave(filter As String, title As String, Optional initdir As String)
  With SaveFileDialog
     .lStructSize = Len(SaveFileDialog)
     .hInstance = App.hInstance
     .lpstrFilter = filter
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = IIf(initdir <> Empty, initdir, CurDir)
     .lpstrTitle = title
     .flags = 0
  End With
End Sub
Public Sub InitOpen(filter As String, title As String, Optional initdir As String)
   filter = Replace(filter, "|", Chr(0))
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hInstance = App.hInstance
     .lpstrFilter = filter
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = IIf(initdir <> Empty, initdir, CurDir)
     .lpstrTitle = title
     .flags = 0
   End With
End Sub
