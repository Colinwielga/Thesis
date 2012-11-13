VERSION 5.00
Begin VB.Form frmVillains 
   BackColor       =   &H00000080&
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   Picture         =   "frmVillains.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox drpVillain 
      Height          =   315
      ItemData        =   "frmVillains.frx":830C
      Left            =   4680
      List            =   "frmVillains.frx":8337
      TabIndex        =   6
      Text            =   "Choose a Villain"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdData 
      BackColor       =   &H008080FF&
      Caption         =   "Display Info"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   3240
      ScaleHeight     =   1755
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   2040
      Width           =   5775
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   255
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H008080FF&
      Caption         =   "Return to Disney Castle"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblVillains 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Disney Villains"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H008080FF&
      Caption         =   "Enter the name of a villain.  "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "frmVillains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Villains
'Lori Nohner
'Written - March 17, 20008
'Objective- the user can select a villain from a drop down box and information about that villain from a text file is displayed in the picture box

Private Sub cmdData_Click()
    Dim Villains2(0 To 13) As String 'declares Villains2 array as string
    Dim CTR As Integer ' Declares counter as integer
   
    picResults.Cls ' clears output box
    CTR = 0 'sets CTR at 0
    Open App.Path & "\Villains2.txt" For Input As #1 'opens villains2 array
        Do Until EOF(1) 'tells to read until the end of file
            Input #1, Villains2(CTR)  'sets data in array as Villains2 (CTR)
            CTR = CTR + 1 'add 1 to the counter to keep track of the length of the array
        Loop 'loops back to the do until end of file
    Close #1 'closes villains.txt file
    
    picResults.Print Villains2(drpVillain.ListIndex) 'matches item in drop down box with data in array
        
    
End Sub

Private Sub cmdExit_Click()
    End 'quits program
End Sub

Private Sub cmdReturn_Click()
    frmVillains.Hide ' hides villains page
    frmDisneyCastle.Show 'returns to Disney home page
End Sub


