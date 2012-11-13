VERSION 5.00
Begin VB.Form Inventory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   Picture         =   "VBProject.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "<==Tables"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Inventory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdAlph 
      Caption         =   "Alphabetize"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   4200
      ScaleHeight     =   6675
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   1800
      Width           =   6255
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Inventory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vinnie Joe's Pub
'Inventory
'Vinnie Schleper, Joey Beltz
'3/26/08
' this form is used to keep track of Inventory.
'   Seeing how much is in stock and searching through long lists of items.
Option Explicit
Dim Food(1 To 50) As String, Quantity(1 To 50) As Integer, CTR As Single
Private OldX As Integer
  Private OldY As Integer
  Private DragMode As Boolean
  Dim MoveMe As Boolean

  Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     MoveMe = True
     OldX = X
     OldY = Y

 End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


     If MoveMe = True Then
         Me.Left = Me.Left + (X - OldX)
         Me.Top = Me.Top + (Y - OldY)
     End If

 End Sub

 Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


     Me.Left = Me.Left + (X - OldX)
     Me.Top = Me.Top + (Y - OldY)
     MoveMe = False

 End Sub

Private Sub cmdAlph_Click()
' this code is used to alphabetize your lists.
Dim Pass As Integer, Pos As Integer, Temp As String, i As Single
picResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Food(Pos) > Food(Pos + 1) Then
            Temp = Food(Pos)
            Food(Pos) = Food(Pos + 1)
            Food(Pos + 1) = Temp
            Temp = Quantity(Pos)               'this keeps the quantity with the item.
            Quantity(Pos) = Quantity(Pos + 1)
            Quantity(Pos + 1) = Temp
        End If
    Next Pos
Next Pass
picResults.Print Tab(30); "Vinnie Joe's Inventory"
    picResults.Print "        "
    picResults.Print "Food Items"; Tab(40); "Quantity"  'tab functions to be neat.
    picResults.Print "-------------------------------------------------------------------------"
For i = 1 To CTR
    picResults.Print Food(i); Tab(40); Quantity(i)
Next i
End Sub

Private Sub cmdBack_Click()
Inventory.Hide
Tables.Show
End Sub

Private Sub cmdRead_Click()
' this code is used to read an text file and put it's contents into 2 arrays.
Dim i As Integer
CTR = 0
Open App.Path & "\Inventory.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Food(CTR), Quantity(CTR)
Loop
    picResults.Print Tab(30); "Vinnie Joe's Inventory"
    picResults.Print "        "
    picResults.Print "Food Items"; Tab(40); "Quantity"
    picResults.Print "-------------------------------------------------------------------------"
For i = 1 To CTR
    picResults.Print Food(i); Tab(40); Quantity(i)
Next i

End Sub

Private Sub cmdSearch_Click()
' this code makes it easier for the employee to find a specific item and how much
'   of it there is easily, especially when there are long item lists.
Dim SearchFor As String
Dim Found As Boolean
Dim i As Integer
Dim Searchfood As String
Searchfood = InputBox("What Item Would You Like To Search For?", "Enter Name")
Found = False
Do While ((Not Found) And (i < CTR))
        i = i + 1
        If Food(i) = Searchfood Then
            Found = True
        End If
    Loop
If Found = True Then
        MsgBox Food(i) & " has " & Quantity(i) & " left in stock."
    Else
        MsgBox Searchfood & " was not found."
    End If

End Sub

  
