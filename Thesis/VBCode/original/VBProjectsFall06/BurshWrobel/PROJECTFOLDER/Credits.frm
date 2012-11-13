VERSION 5.00
Begin VB.Form Form25 
   Caption         =   "Form25"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form25"
   Picture         =   "Credits.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Order Online!"
      Height          =   1455
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "by Laurie Schneider Adams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   8280
      Width           =   3375
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form25
'Bursh,Wrobel
'11-1-06
'This is our Credits Form. Basically all of our information came
'from this book, and with this form, we give the user a chance
'to buy it online, at amazon.com.  What a great book!  This code
'was found online under a template.
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const INFINITE As Long = -1
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_SHOWNORMAL As Long = 1

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
     
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long
         
         

Private Sub Command1_Click()

   Dim sURL As String
   
  'the URL to open, of course!
   sURL = "http://www.amazon.com/History-Western-Core-Concepts-CD-ROM/dp/0072997680/sr=8-1/qid=1162191721/ref=pd_bbs_sr_1/002-1669387-8244008?ie=UTF8&s=books"
   
  'if the call returns false, display a message
   If Not StartNewBrowser(sURL) Then
   
      MsgBox "No dice!"
      
   End If
   
End Sub


Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim Start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   Dim sCmdLine As String
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
   
      sCmdLine = BuildCommandLine(sBrowser)
      
     'prepare STARTUPINFO members
      With Start
         .cb = Len(Start)
         .dwFlags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              sCmdLine & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, Start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function


Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function


Private Function BuildCommandLine(ByVal sBrowser As String) As String

  'just in case the returned string is mixed case
   sBrowser = LCase$(sBrowser)
   
  'try for internet explorer
   If InStr(sBrowser, "iexplore.exe") > 0 Then
      BuildCommandLine = " -nohome "
   
  'try for netscape 4.x
   ElseIf InStr(sBrowser, "netscape.exe") > 0 Then
      BuildCommandLine = " "
   
  'try for netscape 7.x
   ElseIf InStr(sBrowser, "netscp.exe") > 0 Then
      BuildCommandLine = " -url "
   
   Else
   
     'not one of the usual browsers, so
     'either determine the appropriate
     'command line required through testing
     'and adding to ElseIf conditions above,
     'or just return a default 'empty'
     'command line consisting of a space
     '(to separate the exe and command line
     'when CreateProcess assembles the string)
      BuildCommandLine = " "
      
   End If
   
End Function


Private Function TrimNull(item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
      TrimNull = Left$(item, pos - 1)
   Else
      TrimNull = item
   End If
   
End Function


Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(MAX_PATH)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function




Private Sub Command2_Click()
    Form1.Show
    Form25.Hide
End Sub
