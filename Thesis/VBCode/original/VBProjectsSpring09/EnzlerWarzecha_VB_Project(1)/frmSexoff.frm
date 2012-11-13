VERSION 5.00
Begin VB.Form frmSexoff 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   9975
      Left            =   7800
      ScaleHeight     =   9975
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton cmdDuluth 
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdRoch 
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   7200
      Width           =   255
   End
   Begin VB.CommandButton cmdTC 
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to Alphabetize Sex Offenders"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   4320
      Picture         =   "frmSexoff.frx":0000
      Top             =   8520
      Width           =   3465
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Duluth"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Rochester"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Twin Cities"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "St. Joseph"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Image frmSexoff 
      Height          =   7860
      Left            =   240
      Picture         =   "frmSexoff.frx":12B52
      Top             =   480
      Width           =   7350
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmSexoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Age(1 To 100) As Integer, CTR As Integer

Private Sub cmdDuluth_Click()
Dim Duluth(1 To 100) As String

 'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Duluth.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Duluth(CTR), Age(CTR)
        picResults.Print Duluth(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in Duluth is..."; CTR
    Close #1
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRoch_Click()
Dim Roch(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Rochester.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Roch(CTR), Age(CTR)
        picResults.Print Roch(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in Rochester is..."; CTR
     Close #1
End Sub

Private Sub cmdTC_Click()
Dim Cities(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Cities.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Cities(CTR), Age(CTR)
        picResults.Print Cities(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in the Twin Cities is..."; CTR
     Close #1
End Sub

Private Sub Command1_Click()
Dim StJoe(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\st.joe.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, StJoe(CTR), Age(CTR)
        picResults.Print StJoe(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in St. Joseph is..."; CTR
     Close #1
End Sub




Private Sub quit_Click()
End
End Sub

Private Sub return_Click()
frmSexoff.Hide
frmHome.Show
End Sub
