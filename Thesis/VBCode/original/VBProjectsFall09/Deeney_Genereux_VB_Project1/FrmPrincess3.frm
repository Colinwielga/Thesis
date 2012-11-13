VERSION 5.00
Begin VB.Form FrmPrincess3 
   Caption         =   "Princess3"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   Picture         =   "FrmPrincess3.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox Picresults1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4560
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton CmdLocation 
      Caption         =   "Sort By Location in the Mall"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton CmdABC 
      Caption         =   "Sort By ABC Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CommandButton CmdStores 
      Caption         =   "Stores in the mall the princess might be interested in going to"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.PictureBox Picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   4560
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Help the princess sort out what store she should go into."
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "The princess needs to give her brain a rest from all that math! It's time to take another trip to the mall!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "FrmPrincess3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form takes two arrays and sorts them into two categories: Stores & Location
'There is a button that sorts the Arrays into ABC order
'There is another button that sorts the arrays into location number order
'The form ends there to go to the next one

    Dim Stores(1 To 7) As String, location(1 To 7) As Integer, Ctr As Integer, tempname As String, templocation As Integer, pass As Integer, pos As Integer, J As Integer
    
    
  
Private Sub CmdABC_Click()
'Puts the array into ABC order

CmdABC.Enabled = False 'makes the ABC button so that you cannot click on it
CmdLocation.Enabled = True 'Makes this button visible


For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Stores(pos) > Stores(pos + 1) Then 'This gets the Stores into ABC order
            tempname = Stores(pos)
            Stores(pos) = Stores(pos + 1)
            Stores(pos + 1) = tempname
            templocation = location(pos) 'This puts the location with the given store
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
            
            
        End If
    Next pos
Next pass
 

    picResults1.Print "Store", "Store Location"
    picResults1.Print "***********************************************************"
    
    For J = 1 To Ctr
             picResults1.Print Stores(J); Tab(20); location(J) 'Prints the results in the picture box
    Next J
    

End Sub


Private Sub CmdLocation_Click()

'Puts the array into Location order

    picResults1.Cls
    CmdLocation.Enabled = False
    

For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If location(pos) > location(pos + 1) Then 'This makes the location with the lowest number first
            templocation = location(pos)
            location(pos) = location(pos + 1)
            location(pos + 1) = templocation
            tempname = Stores(pos) 'Makes the store name match up with the location
            Stores(pos) = Stores(pos + 1)
            Stores(pos + 1) = tempname
            
            
        End If
    Next pos
Next pass
 

    picResults1.Print "Store", "Store Location"
    picResults1.Print "***********************************************************"
    
    For J = 1 To Ctr
             picResults1.Print Stores(J); Tab(20); location(J) 'Prints the results in a picture box
    Next J
    cmdNext.Enabled = True
End Sub

Private Sub CmdNext_Click()
'Goes to the next form
    FrmPrincess3.Hide 'Hides this form and opens up the next
    FrmPrincessEnd.Show
    CmdABC.Enabled = False
    CmdLocation.Enabled = False
    cmdNext.Enabled = False
    CmdStores.Enabled = True
    PicResults.Cls
    picResults1.Cls
    
    
End Sub

Private Sub cmdQuit_Click()
'Ends the program
    End
End Sub

Private Sub CmdStores_Click()
    
    Ctr = 0
    
    Open App.Path & "\Stores.txt" For Input As #2
    
    PicResults.Print "Stores", "Location in the Mall"
    PicResults.Print "***************************************************"
    
    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        
        Ctr = Ctr + 1
        'Read next data set from the file into the array
        'and print the data
        
        Input #2, Stores(Ctr), location(Ctr) 'gets input from the file
        
        
        PicResults.Print Stores(Ctr); Tab(20); location(Ctr) 'Prints the needed information in the picture box
        
        Loop 'Loops until everything is printed throughout the whole file
        Close #2
        PicResults.Print "*******************************************************"
        
        CmdStores.Enabled = False
        CmdABC.Enabled = True 'Enables the ABC button, disables the other two
        CmdLocation = False
        
    
    
    
End Sub
