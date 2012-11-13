VERSION 5.00
Begin VB.Form frmBegin 
   BackColor       =   &H80000012&
   Caption         =   "Read File"
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12285
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H80000014&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      MaskColor       =   &H00C00000&
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtTheName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   4800
      Width           =   4575
   End
   Begin VB.CommandButton cmdReadFile 
      BackColor       =   &H80000014&
      Caption         =   "Read File"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5760
      MaskColor       =   &H80000014&
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReadFile_Click()
'This part of the program makes it so that when you click the
'Command Button that says "read file", I file called MyNames
'Will be read into an array and then the file is closed.

    'Starting the counter of the file at 0 or the beginning.
    Ctr = 0
    'Opening the Name file as file #1 and as an input file.
    Open App.Path & "\MyNames.txt" For Input As #1
    'The Do While begins reading the file and keeps on looping until it is complete
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, MyName(Ctr)
    Loop
    'Closing the file
    Close #1
    'A message box will appear on the screen telling you to enter your name by clicking "ok"
    MsgBox "Enter Your Name", , "Enter Name"
End Sub
Private Sub cmdSubmit_Click()
'This piece switches to the second form but only if the name that is entered into the
'Text box "theName" can be found within the name file from before.
'Otherwise the form stays the same until a correct name is entered.
    
    'Here the variable "theName" is being declared
    Dim theName As String
    'Found is false meaning that the program has not found the name within the file yet.
    Found = False
    'This is telling us the "theName" will be whatever is entered into the text box.
    theName = txtTheName.Text
    'This is the beginning of a loop and it will read through the seven names that are in the file.
    For I = 1 To 7
        'This is a nested if and it compares what was typed into the text box with the information in
        'the name file. If there is a match then found becomes true
        If StrComp(theName, MyName(I), vbTextCompare) = 0 Then
            Found = True
            Pos = I
            'This changes from the first form to the second form while keeping the third form from coming up.
            frmBegin.Visible = False
            frmNameSelect.Visible = True
            frmShopping.Visible = False
        End If
    Next I
End Sub


