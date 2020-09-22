Attribute VB_Name = "modMain"
Option Explicit
'Created by Dave Johnston, Oct 23,2001
      
Private Declare Function GetVersionExA Lib "kernel32" _
  (lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
   "GetWindowsDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Const mcsTitle As String = "Running Batch File"

Private Function GetOSVersion() As String
  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer

  On Error GoTo ErrorHandler
  
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   With osinfo
   Select Case .dwPlatformId
      Case 1
         If .dwMinorVersion = 0 Then
            GetOSVersion = "Windows 95"
         ElseIf .dwMinorVersion = 10 Then
            GetOSVersion = "Windows 98"
         End If
      Case 2
         If .dwMajorVersion = 3 Then
            GetOSVersion = "Windows NT 3.51"
         ElseIf .dwMajorVersion = 4 Then
            GetOSVersion = "Windows NT 4.0"
         ElseIf .dwMajorVersion = 5 Then
            GetOSVersion = "Windows 2000"
         End If
      Case Else
         GetOSVersion = "Failed"
   End Select
   End With

NormalExit:
  Exit Function

ErrorHandler:
  MsgBox "I'm sorry, an error has occured and I don't know how to deal with it. Procedure: GetOSVersion  Error: " & Err.Number & " - " & Err.Description, vbExclamation, mcsTitle
  Exit Function

End Function

Private Function GetSystemDir() As String
  Dim mcsTitle As String
 
  Dim Buffer As String
  Dim rc As Long
  
  On Error GoTo ErrorHandler

  Buffer = Space$(256)
  rc = GetSystemDirectory(Buffer, Len(Buffer))
  GetSystemDir = LCase$(Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1))
 
NormalExit:
  Exit Function

ErrorHandler:
  MsgBox "I'm sorry, an error has occured and I don't know how to deal with it. Procedure: GetSystemDir  Error: " & Err.Number & " - " & Err.Description, vbExclamation, mcsTitle
  Exit Function

End Function

Private Function GetWindowsDir() As String
 
  Dim Buffer As String
  Dim rc As Long
  
  On Error GoTo ErrorHandler

  Buffer = Space$(256)
  rc = GetWindowsDirectory(Buffer, Len(Buffer))
  GetWindowsDir = LCase$(Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1))
 
NormalExit:
  Exit Function

ErrorHandler:
  MsgBox "I'm sorry, an error has occured and I don't know how to deal with it. Procedure: GetWindowsDir  Error: " & Err.Number & " - " & Err.Description, vbExclamation, mcsTitle
  Exit Function

End Function

Private Sub Main()

  Dim sMsg As String
  Dim sCmdLine As String
  Dim iCmdLineLen As Integer
  Dim sBatchFileName As String
  Dim vWindowOption As Variant
  Dim vCloseOption As Variant
  Dim iWindowOption As Integer
  Dim sCloseOption As String
  Dim iBPos As Integer
  Dim iWPos As Integer
  Dim iCPos As Integer
  Dim sOSVersion As String
  Dim sCommandProcessor As String
  
  On Error GoTo ErrorHandler

  sCmdLine = UCase(Command())
  iCmdLineLen = Len(sCmdLine)
  
  sMsg = "Usage: RunBat /B <Batch file name and path> /W <0-4,6> /C<YN>" & vbCrLf & vbCrLf & _
      "Where: /W is window option 0=Hide, 1=Normal, 2=Minimized, 3=Maximized, 4=Normal (No Focus), 6=Minimized (No Focus), /C is close option Y=Close after finished" & vbCrLf & vbCrLf & _
      "Example: RunBat /B C:\temp\mybat.bat /W 1 /C"
  
  If iCmdLineLen = 0 Then
    MsgBox "No parameters specified!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
    GoTo NormalExit
  End If

  iBPos = InStr(sCmdLine, "/B")
  iWPos = InStr(sCmdLine, "/W")
  iCPos = InStr(sCmdLine, "/C")
  
  If iBPos = 0 Or iWPos = 0 Or iCPos = 0 Then
    MsgBox "Missing parameters!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
    GoTo NormalExit
  End If
           
  sBatchFileName = Trim(Mid(sCmdLine, iBPos + 2, iWPos - (iBPos + 2)))
  vWindowOption = Trim(Mid(sCmdLine, iWPos + 2, iCPos - (iWPos + 2)))
  vCloseOption = Trim(Mid(sCmdLine, iCPos + 2))
             
  If sBatchFileName = "" Or vWindowOption = "" Or vCloseOption = "" Then
    MsgBox "Missing parameters!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
    GoTo NormalExit
  End If
           
  Select Case vWindowOption
    Case "0": iWindowOption = vbHide
    Case "1": iWindowOption = vbNormalFocus
    Case "2": iWindowOption = vbMinimizedFocus
    Case "3": iWindowOption = vbMaximizedFocus
    Case "4": iWindowOption = vbNormalNoFocus
    Case "6": iWindowOption = vbMinimizedNoFocus
    Case Else
      MsgBox "Invalid Window Option!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
      GoTo NormalExit
  End Select
  
  Select Case vCloseOption
    Case "Y": sCloseOption = "/C"
    Case "N": sCloseOption = "/K"
    Case Else
      MsgBox "Invalid Close Option!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
      GoTo NormalExit
  End Select
    
  If Dir(sBatchFileName) = "" Then
    MsgBox "Batch file does not exist!" & vbCrLf & vbCrLf & sMsg, vbExclamation, mcsTitle
    GoTo NormalExit
  End If
  
  sOSVersion = GetOSVersion
  
  Select Case sOSVersion
    Case "Windows 95", "Windows 98": sCommandProcessor = GetWindowsDir & "\" & "Command.com"
    Case "Windows NT 3.51", "Windows NT 4.0", "Windows 2000": sCommandProcessor = GetSystemDir & "\" & "Cmd.exe"
    Case Else
      sCommandProcessor = GetWindowsDir & "\" & "Command.com"
  End Select
  
'  MsgBox "About to execute: " & sCommandProcessor & " " & sCloseOption & " """ & sBatchFileName & """", iWindowOption
  Shell sCommandProcessor & " " & sCloseOption & " """ & sBatchFileName & """", iWindowOption
  
NormalExit:
  Exit Sub

ErrorHandler:
  MsgBox "I'm sorry, an error has occured and I don't know how to deal with it. Procedure: Main  Error: " & Err.Number & " - " & Err.Description, vbExclamation, mcsTitle
  Exit Sub
  
End Sub

