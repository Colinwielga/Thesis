VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "The Record -- Display ad -- Black & white -- Size"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form3"
   ScaleHeight     =   5775
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Continue 
      Caption         =   "Click here to continue the process of placing an ad."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   4575
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   5160
      ScaleHeight     =   2115
      ScaleWidth      =   4395
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "blackandwhite.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "blackandwhite.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "blackandwhite.frx":1CBE4
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Custom 
      Caption         =   "Click to customize the size of your ad."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Preset 
      Caption         =   "Click to choose pre-set size"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"blackandwhite.frx":24156
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   9495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"blackandwhite.frx":242DC
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1/4 page - $100 (save $20)    1/2 page - $200 (save $30)    3/4 page - $300 (save $40)    Full page - $350 (save $50)"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 3(blackandwhite.frm), Kristen Nowak, 3-14-04, The purpose of this form is to allow the user to select the size of a black and white display ad.

Private Sub Continue_Click()
Form3.Hide 'continue to the next part of the program, so that the user can select from other options
Form2.Show
End Sub

Private Sub Custom_Click()
Dim myTable(1 To 20, 1 To 7) As Integer
Dim row As Integer
Dim column As Integer
Dim Inch As Integer
Dim col As Integer
Dim Path As String

Path = "N:\CS130\handin\Nowak, Kristen\"

Open Path & "black.txt" For Input As #1
'read data into a two-dimensional array
For row = 1 To 20
    For column = 1 To 7
        Input #1, myTable(row, column)
    Next column
Next row
Close


row = InputBox("Enter the height of the ad in column inches (1 to 20)") 'user determines height
column = InputBox("Enter the width of the ad in number of columns (1 to 7)") 'user determines width

Total = myTable(row, column) 'determines total by matching up row and column in the table
Results.Print "The cost of an ad"; row; "column inches by"; column; "columns is "; FormatCurrency(Total) 'prints the cost of an ad

Custom.Enabled = False
Preset.Enabled = False
Continue.Enabled = True
End Sub
Private Sub Form_Load()
Continue.Enabled = False 'requires user to select a size before going on to the next page
End Sub

Private Sub Preset_Click()
Size = InputBox("You have chosen a pre-set discount size. What size would you like? Type 1 for 1/4 page, 2 for 1/2 page, 3 for 3/4 page and 4 for a full page.")
If Size = 1 Then 'calculates and shows the cost of a a 1/4 page ad
    Total = 100
    Results.Print "A 1/4 page B&W ad will cost "; FormatCurrency(Total); " for one week."
    Custom.Enabled = False
    Preset.Enabled = False
    Continue.Enabled = True
ElseIf Size = 2 Then 'calculates and shows the cost of a 1/2 page ad
    Total = 200
    Results.Print "A 1/2 page B&W ad will cost "; FormatCurrency(Total); " for one week."
    Custom.Enabled = False
    Preset.Enabled = False
    Continue.Enabled = True
ElseIf Size = 3 Then 'calculates and shows the cost of a 3/4 page ad
    Total = 300
    Results.Print "A 3/4 page B&W ad will cost "; FormatCurrency(Total); " for one week."
    Custom.Enabled = False
    Preset.Enabled = False
    Continue.Enabled = True
ElseIf Size = 4 Then 'calculates and shows the cost of a full page ad
    Total = 350
    Results.Print "A full page B&W ad will cost "; FormatCurrency(Total); " for one week."
    Custom.Enabled = False
    Preset.Enabled = False
    Continue.Enabled = True
Else: MsgBox "Error: You entered in an incorrect value. Please try again." 'user did not enter a number between 1 and 4

End If

End Sub


Private Sub Quitbutton_Click()
End
End Sub

