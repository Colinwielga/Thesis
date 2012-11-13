VERSION 5.00
Begin VB.Form frmRoomSize 
   Caption         =   "Room Size Selection Page"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   Picture         =   "frmRoomSize.frx":0000
   ScaleHeight     =   9285
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToCheckin 
      Caption         =   "Go Back to Check-in Page"
      Height          =   1695
      Left            =   10080
      TabIndex        =   12
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton CmdMasterSuite 
      Caption         =   "Master Suite"
      Height          =   1695
      Left            =   6360
      TabIndex        =   4
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdSmallSuite 
      Caption         =   "Small Suite"
      Height          =   1695
      Left            =   6360
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdKing 
      Caption         =   "One King Bed"
      Height          =   1695
      Left            =   6360
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton CmdQueenBed 
      Caption         =   "One Queen Bed"
      Height          =   1695
      Left            =   6360
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton CmdDouble 
      Caption         =   "Two Double Beds"
      Height          =   1695
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTelevision 
      Caption         =   "4.) Television with HBO© included"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label lblBathroom 
      Caption         =   "3.) Bathroom with Shower, Sink, and Toilet"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   6720
      Width           =   3735
   End
   Begin VB.Label lblKitchen 
      Caption         =   "2.) Small Kitchen"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   6480
      Width           =   3735
   End
   Begin VB.Label lblmarks 
      Caption         =   "**************************************************************"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Label lblFridge 
      Caption         =   "1.) Small Refridgerator"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   6240
      Width           =   3735
   End
   Begin VB.Label lblRoomDescriptions 
      Alignment       =   2  'Center
      Caption         =   "All our rooms come with these things standard:"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   6600
      Width           =   615
   End
End
Attribute VB_Name = "frmRoomSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Room Size Selection
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   This is the page where we find out what size room our guest would
'           like to stay in. We give them the information about each in
'           messageboxes and also ask them to input whether or not they would
'           like a smoking room.
    


Option Explicit
Private Sub cmdBackToCheckin_Click()
'hides the room size menu, and shows the Main menu
    frmRoomSize.Hide
    frmMainMenu.Show
End Sub

Private Sub CmdDouble_Click()
    MsgBox "A Double room is a standard room at our hotel with 2 Double beds.", , "Double"
    
'THIS IS FOR ALL THE BUTTONS FOR DIFFERENT SIZED ROOMS!!!
'********************************************************
'when clicked, a messagebox appears telling what the user what the room offers.
    
'sets the room in the array according to which room was selected
    SelectedRoom = Rooms(1)
    SelectedPrice = Price(1)
    
'asks the user if they want a smoking room
    Smoking = InputBox("Would you like a smoking room? (yes or no)", "Smoking", "")
    
'sets cozy... which is a public string to be known as Double (or whatever the
'room size is)
    Cozy = "Double"

'If smoking is selected, then the smoking form is shown
    If Smoking = "yes" Then
        frmStayPrice.Show
        frmRoomSize.Hide
'Otherwise, the non-smoking form is shown
    ElseIf Smoking = "no" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    Else
        MsgBox "Please input yes or no.", , "Error"
    End If
    
    
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub cmdKing_Click()
    MsgBox "A King room is a standard room at our hotel with 1 King bed.", , "King"
        
        
    SelectedRoom = Rooms(3)
    SelectedPrice = Price(3)
    
    Smoking = InputBox("Would you like a smoking room? (yes or no)", "Smoking", "")
    
    Cozy = "King"
    
    If Smoking = "yes" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    ElseIf Smoking = "no" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    Else
        MsgBox "Please input yes or no.", , "Error"
    End If
    
    
    
End Sub

Private Sub CmdMasterSuite_Click()
    MsgBox "The Master Suite is the finest room in the Hotel. 2 King beds in seaprate rooms, a full kitchen with chef, a living room and dining room, and your own Jacuzzi©.", , "Master-Suite"
    
    
    
    SelectedRoom = Rooms(5)
    SelectedPrice = Price(5)
    
    Smoking = InputBox("Would you like a smoking room? (yes or no)", "Smoking", "")
    
    Cozy = "MasterSuite"
    
    If Smoking = "yes" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    ElseIf Smoking = "no" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    Else
        MsgBox "Please input yes or no.", , "Error"
    End If
  
    
    
End Sub

Private Sub CmdQueenBed_Click()
    MsgBox "A Queen room is a standard room at our hotel with 1 Queen bed.", , "Queen"
    
    
    SelectedRoom = Rooms(2)
    SelectedPrice = Price(2)
    
    Smoking = InputBox("Would you like a smoking room? (yes or no)", "Smoking", "")
    
    Cozy = "Queen"
    
    If Smoking = "yes" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    ElseIf Smoking = "no" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    Else
        MsgBox "Please input yes or no.", , "Error"
    End If
    
    
    
    
End Sub

Private Sub cmdSmallSuite_Click()
    MsgBox "In our Small Suite we offer 1 King bed, a full kitchen, and a separate living room with Couch, table, and chairs.", , "Small Suite"
    
    
    
    SelectedRoom = Rooms(4)
    SelectedPrice = Price(4)
   
    Smoking = InputBox("Would you like a smoking room? (yes or no)", "Smoking", "")
    
    Cozy = "SmallSuite"
    
    If Smoking = "yes" Then
        frmStayPrice.Show
        frmRoomSize.Hide
    ElseIf Smoking = "no" Then
        
        frmStayPrice.Show
        frmRoomSize.Hide
    Else
        MsgBox "Please input yes or no.", , "Error"
    End If
    
    
    
End Sub

Private Sub Form_Load()
    Dim CTR As Integer
    
'When the form initially loads, the "roomsandrates" text file gets opened, and
'the information within gets written into arrays. we used that to price out our
'rooms as well as keep the rooms and prices together.
    Open App.Path & "\roomsandrates.txt" For Input As #1
        
        CTR = 1
        
        Do While Not EOF(1)
            Input #1, Rooms(CTR), Price(CTR)
            CTR = CTR + 1
        Loop
    Close #1

    
    
End Sub
