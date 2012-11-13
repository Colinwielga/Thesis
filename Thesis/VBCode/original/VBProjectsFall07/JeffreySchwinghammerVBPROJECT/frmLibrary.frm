VERSION 5.00
Begin VB.Form frmLibrary 
   BackColor       =   &H80000007&
   Caption         =   "Office Room"
   ClientHeight    =   8130
   ClientLeft      =   225
   ClientTop       =   1410
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   9254.032
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   360
      Picture         =   "frmLibrary.frx":0000
      ScaleHeight     =   6375
      ScaleWidth      =   10695
      TabIndex        =   4
      Top             =   120
      Width           =   10695
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Click to View Selected Item"
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
   End
   Begin VB.ListBox lstchecklist 
      Height          =   645
      Left            =   9480
      TabIndex        =   2
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturntoHub 
      Caption         =   "Return to Previous Room"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin VB.PictureBox picLibraryTxt 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   6720
      Width           =   6015
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
    If lstchecklist = "" Then
        picLibraryTxt.Cls
        
        picLibraryTxt.Print "Please select from the list what you want to view."
    End If
    
    If lstchecklist = "Book Case" Then
        frmBookCase.Show
        frmLibrary.Hide
    End If
    
    If lstchecklist = "Portraits" Then
        frmPortraits.Show
        frmLibrary.Hide
    End If
    
    If lstchecklist = "Desk" Then
        picLibraryTxt.Cls
        picLibraryTxt.Print "There is alot of papers and books on the desk."
    End If
    
End Sub

Private Sub cmdReturntoHub_Click()
    picLibraryTxt.Cls
    If Gun = False Or EmblemTwo = False Then    ' Player can not leave till he succeeds in this room
        picLibraryTxt.Print "The door won't budge. There might be something in here that "
        picLibraryTxt.Print "will open the door."
    Else
        frmHub.Visible = True
        frmLibrary.Visible = False
    End If
End Sub

Private Sub Form_activate()
     
   Dim answer As Integer
     
     picLibraryTxt.Cls
     'Introduction to room at first entrance
    If LeftRoomCheck = False Then
        picLibraryTxt.Print "After you step through the door, it slams shut and locks."
        picLibraryTxt.Print "'Not again', you mutter to yourself. You look around the room and notice"
        picLibraryTxt.Print "that there is a book shelf, some odd portraits, and a desk."
        
              'Filling up Check List
        lstchecklist.AddItem ("Book Case")
        lstchecklist.AddItem ("Portraits")
        lstchecklist.AddItem ("Desk")
         
        LeftRoomCheck = True
    Else
        If Gun = True And EmblemTwo = True Then
            answer = MsgBox("You hear the door unlock. You can leave this room now.", vbOKOnly)
        End If
    End If
    
End Sub
