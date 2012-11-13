VERSION 5.00
Begin VB.Form frmfans 
   BackColor       =   &H80000003&
   Caption         =   "Fans"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdwebsite 
      Caption         =   "Offical Counting Crows Website"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdfindage 
      Caption         =   "Click here to find fans your Age"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdfindstate 
      Caption         =   "Click here to find fans in your State"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdlist 
      Caption         =   "Click here to see a list of CC Fans"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   5880
      ScaleHeight     =   4515
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back to the Main Page"
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Matt Proulx"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
End
Attribute VB_Name = "frmfans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : CountingCrows (Matt Proulx's VB Project.vbp)
'Form Name : frmfans (frmfans.frm)
'Author: Matt Proulx
'Date Written: March 13, 2004
'Purpose of the Form: 'This form will let the user see a list of registered Counting Crows fans. It will also let the
                      'user input a state or an age of a fan and it will then tell the user if there is anyone in the
                      'list the same as your inputted state or age.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Dim people(1 To 10) As String
Dim state(1 To 10) As String
Dim age(1 To 10) As Single
Dim CTR As Integer
Private Declare Function ShellExecute _
    Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long
Private Sub cmdback_Click() 'Let the user return to the main page
    frmfans.Hide
    frmtitle.Show
End Sub
Private Sub cmdfindage_Click()
Dim theirage As Single
Dim found As Boolean
Dim placeCtr As Integer 'keeps track of where you are in the list
placeCtr = 0
found = False
'keeps looking as long as you have not found what you are looking for
'and you have not reached the end of the array
theirage = InputBox("Enter an age")
CTR = 10
Do While (Not found) And (placeCtr < CTR) 'Searches through the array
    placeCtr = placeCtr + 1
    If age(placeCtr) = theirage Then
        found = True
        picResults.Print "********************************************"
        picResults.Print people(placeCtr); " "; "is the same age as your input"
    End If
Loop
If Not found Then
    MsgBox "Sorry, but no one in our registery is that age", , "Age not found"
End If
End Sub
Private Sub cmdfindstate_Click()
Dim theirstate As String
Dim found As Boolean
Dim placeCtr As Integer 'keeps track of where you are in the list
placeCtr = 0
found = False
'keeps looking as long as you have not found what you are looking for and
'you have not reached the end of the array
theirstate = InputBox("Enter a state")
CTR = 10
Do While (Not found) And (placeCtr < CTR)
    placeCtr = placeCtr + 1
    If state(placeCtr) = theirstate Then
        found = True
        picResults.Print "********************************************"
        picResults.Print people(placeCtr); " "; "lives in your state!"
    End If
Loop
If Not found Then
    MsgBox "Sorry, but no one in our registery is from that state", , "State not found"
End If
End Sub
Private Sub cmdlist_Click()
picResults.Cls 'Clears the picture window
CTR = 0
Open Path & "names.txt" For Input As #1
picResults.Print "These are registered Counting Crows Fans"
picResults.Print "*******************************************************************"
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, people(CTR), state(CTR), age(CTR)
        picResults.Print people(CTR)
    Loop
    Close
    cmdfindstate.Enabled = True 'Now that the array has been loaded, the user can search for age and state of fans
    cmdfindage.Enabled = True
End Sub
Private Sub CmdWebsite_Click()
'Brings the user to the offical Counting Crows webpage
    Dim r As Long
    r = ShellExecute(0, "open", "http://www.countingcrows.com", 0, 0, 1)
End Sub
Private Sub Form_Load()
    Path = "N:\CS130\handin\Proulx, Matt\" 'Tells the computer where to find the information (list)
    cmdfindstate.Enabled = False 'Both find state and find age can not be clicked until the array has been loaded
    cmdfindage.Enabled = False
End Sub
