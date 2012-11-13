VERSION 5.00
Begin VB.Form Equipment 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrout 
      BackColor       =   &H00FFFF00&
      Caption         =   "Let's Find Out More About Minnesota Trout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdTerm 
      BackColor       =   &H00FFFF00&
      Caption         =   "Is There A Term You Don't Understand?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2535
   End
   Begin VB.PictureBox picPicture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   5040
      ScaleHeight     =   6075
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   960
      Width           =   4815
   End
   Begin VB.CommandButton cmdFly 
      BackColor       =   &H00FFFF80&
      Caption         =   "What Kind Of Equipment Will I need?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is entitled Equipment
'It was written by Kevin Lundeen
'March 22nd, 2009
'
'The purpose of this form is to give the audience an idea of the equipment that they can expect to need.
'It also gives a brief description of the equipment if they are unaware of what it is.
'
'Declare all variables
    
    Dim Equipment(1 To 30) As String
    Dim Amount(1 To 30) As Integer, CTR As Integer
    Dim TempEquip As String, TempAmount As Integer
    Dim Pass As Integer, Pos As Integer
    Dim I As Integer
    
Private Sub cmdFly_Click()

'This subroutine reads information from a data file, stores it, and sorts it in parallel arrays
        
    Open App.Path & "\FlyEquipment.txt" For Input As #1     'Opens up the file
    
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Equipment(CTR), Amount(CTR)               'Reads the data and then stores it
    Loop
    
    picPicture2.Cls
    picPicture2.Print "Equipment", "Amount"
    picPicture2.Print "*****************************"
    
    For Pass = 1 To CTR - 1                                 'The bubble sort sorts the data by the amount needed
        For Pos = 1 To Pass - 1
            If Amount(Pos) > Amount(Pos + 1) Then
                TempAmount = Amount(Pos)                    'These steps hold the data into temporary holding cells and sort them out
                Amount(Pos) = Amount(Pos + 1)               'by size.
                Amount(Pos + 1) = TempAmount                '
                TempEquip = Equipment(Pos)                  '
                Equipment(Pos) = Equipment(Pos + 1)         '
                Equipment(Pos + 1) = TempEquip              '
            End If
        Next Pos
    Next Pass
    
    For I = 1 To CTR
        picPicture2.Print Equipment(I), Amount(I)           'Displays the sorted data in parallel arrays
    Next I
    
   
        
End Sub

Private Sub cmdQuit2_Click()
    End     'Ends the program
End Sub

Private Sub cmdTerm_Click()

'This subroutine gives a brief definition to the user if they are unaware of what a paticular item is used for.

    Dim Found As Boolean                'Declare variables
    Dim NameOfEquipment As String       '
    
    picPicture2.Cls
    
    Found = False
    
    'Gets the item that needs to be clarified from an input box
    NameOfEquipment = InputBox("What term do you not understand in the list? Please enter it exactly as you see it in the Picture Box.")
    
    'The ElseIf loops makes connecting the searched for word with its definition easy.
    'If the word is found the Found variable is set to true, which will not then display a message about the word not being found.
            If NameOfEquipment = "Fly Rod" Then
                picPicture2.Print "A fly rod is used in fishing to" & vbCrLf & " cast the line out to the fish."
                Found = True
                ElseIf NameOfEquipment = "Reel" Then
                    picPicture2.Print "The reel is what the line is wrapped" & vbCrLf & " around and stored in while fishing."
                    Found = True
                ElseIf NameOfEquipment = "Dry Flies" Then
                    picPicture2.Print "Dry flies are artificial bait that float" & vbCrLf & " on the surface of the water."
                    Found = True
                ElseIf NameOfEquipment = "Nymphs" Then
                    picPicture2.Print "Nymphs are baits which sink and imitate a" & vbCrLf & " particular insects larvae stage."
                    Found = True
                ElseIf NameOfEquipment = "Creel" Then
                    picPicture2.Print "A creel is a basket or pouch that you" & vbCrLf & " put fish into that you wish to keep."
                    Found = True
                ElseIf NameOfEquipment = "Waders" Then
                    picPicture2.Print "Waders are waterproof overalls that keep" & vbCrLf & " you dry while wading through the stream."
                    Found = True
            End If
    
    
    If (Not Found) Then
        MsgBox ("That word wasn't in the array. Please enter the term exactly as you see it in the picture box.")
    End If
               
    
End Sub

Private Sub cmdTrout_Click()
'This subroutine shows the form "Trout"
    Trout.Show

End Sub
