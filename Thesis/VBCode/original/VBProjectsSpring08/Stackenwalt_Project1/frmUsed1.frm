VERSION 5.00
Begin VB.Form frmUsed1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   1050
   ClientTop       =   1230
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13335
   Begin VB.CommandButton cmdSearch3 
      Caption         =   "Search"
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch2 
      Caption         =   "Search"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CheckBox chknum5 
      Caption         =   "Check5"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chknum4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chknum3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chknum2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chknum1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtType 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   3615
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00808080&
      Caption         =   "Go Back to Main Page"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   7695
      Left            =   3960
      ScaleHeight     =   7635
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   840
      Width           =   9135
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H000000FF&
      Caption         =   "List all cars and prices."
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      Caption         =   "$7,501 and higher"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      Caption         =   "$5,001- $7,500"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Caption         =   "$3,001- $5,000"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "$1,501- $3,000"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "$0- $1,500"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblModels 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "(SUV, Car, Truck)"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Label lblColors 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "(Red, Blue, Yellow, Black, Silver, White)"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label lblModel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Search By Type"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Search By Color"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Search By Price"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmUsed1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name- Stack's Car Lot
'Form Name- frmUsed1
'Author- Nick Stackenwalt
'Date Written- Saturday March 09, 2008
'Objective- This form is used to search through all of our used cars
'Other comments- The user can look at the entire list of our cars, and can sort through them
                'By price, color, or type.
Dim color(1 To 100) As String
Dim mileage(1 To 100) As String
Dim price(1 To 100) As String
Dim Model(1 To 100) As String
Dim VeType(1 To 100) As String
Dim pass As Integer, pos As Integer
Dim ctr As Integer
Dim J As Integer

Private Sub chknum1_Click()
If chknum1.Value = 1 Then      'Checks the first box
    chknum2.Value = 0      'Unchecks any other boxes that were checked
    chknum3.Value = 0
    chknum4.Value = 0
    chknum5.Value = 0
End If
End Sub

Private Sub chknum2_Click()
If chknum2.Value = 1 Then      'Checks the second box
    chknum1.Value = 0      'Unchecks any other boxes that were checked
    chknum3.Value = 0
    chknum4.Value = 0
    chknum5.Value = 0
End If
End Sub

Private Sub chknum3_Click()
If chknum3.Value = 1 Then      'Checks the third box
    chknum2.Value = 0      'Unchecks any other boxes that were checked
    chknum1.Value = 0
    chknum4.Value = 0
    chknum5.Value = 0
End If
End Sub

Private Sub chknum4_Click()
If chknum4.Value = 1 Then      'Checks the fourth box
    chknum2.Value = 0      'Unchecks any other boxes that were checked
    chknum3.Value = 0
    chknum1.Value = 0
    chknum5.Value = 0
End If
End Sub

Private Sub chknum5_Click()
If chknum5.Value = 1 Then      'Checks the fifth box
    chknum2.Value = 0      'Unchecks any other boxes that were checked
    chknum3.Value = 0
    chknum4.Value = 0
    chknum1.Value = 0
End If
End Sub

Private Sub cmdAll_Click()
ctr = 0
picResults.Cls       'Clears the Models screen
picResults.Print "Model and Year                                      Type                                        Color                               Milage               Price"     'Prints "Type Model and Year   Color  Milage  Price"
picResults.Print "********************************************************************************************************************************************************"      'Prints "*********************"
Open App.Path & "\listall.txt" For Input As #1      'Opens list of all used cars in seperate arrays
    Do While Not EOF(1)     'Tells it to read entire file
        ctr = ctr + 1
        Input #1, Model(ctr), VeType(ctr), color(ctr), mileage(ctr), price(ctr)      'Puts everything into 4 seperate arrays
    Loop
    For J = 1 To ctr        'Prints out the list
        picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
    Next J
Close #1        'Closes the file
End Sub

Private Sub cmdMain_Click()
frmUsed1.Hide      'Hides Used cars form
frmMain.Show        'Shows the main form
End Sub

Private Sub cmdSearch_Click()
picResults.Cls      'Clears the picResults box
picResults.Print "Model and Year                                        Type                                      Color                               Milage               Price"     'Prints "Type Model and Year   Color  Milage  Price"
picResults.Print "***********************************************************************************************************************************"      'Prints "*********************"
Dim Vcolor As String        'Dims car color variable
Vcolor = txtColor       'Says that what you put in text box is the color to search for
For J = 1 To ctr        'Finds color in list and prints all matches
    If color(J) = txtColor Then
        picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
    End If
Next J
End Sub

Private Sub cmdSearch2_Click()
picResults.Cls      'Clears the picResults box
picResults.Print "Model and Year                                        Type                                      Color                               Milage               Price"     'Prints "Type Model and Year   Color  Milage  Price"
picResults.Print "****************************************************************************************************************************************"      'Prints "*********************"
Dim Vtype As String        'Dims car type variable
Vtype = txtType       'Says that what you put in text box is the type to search for
For J = 1 To ctr        'Finds car type in list and prints all matches
    If VeType(J) = txtType Then
        picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
    End If
Next J
End Sub

Private Sub cmdSearch3_Click()
picResults.Cls      'Clears the picResults box
picResults.Print "Model and Year                                        Type                                      Color                               Milage               Price"     'Prints "Type Model and Year   Color  Milage  Price"
picResults.Print "****************************************************************************************************************************************"      'Prints "*********************"
If chknum1.Value = 1 Then
    For J = 1 To ctr        'Finds car price in list and prints all matches
        If price(J) > 0 And price(J) <= 1500 Then
            picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
        End If
    Next J
End If
If chknum2.Value = 1 Then
    For J = 1 To ctr        'Finds car price in list and prints all matches
        If price(J) > 1501 And price(J) <= 3000 Then
            picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
        End If
    Next J
End If
If chknum3.Value = 1 Then
    For J = 1 To ctr        'Finds car price in list and prints all matches
        If price(J) > 3001 And price(J) <= 5000 Then
            picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
        End If
    Next J
End If
If chknum4.Value = 1 Then
    For J = 1 To ctr        'Finds car price in list and prints all matches
        If price(J) > 5001 And price(J) <= 7500 Then
            picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
        End If
    Next J
End If
If chknum5.Value = 1 Then
    For J = 1 To ctr        'Finds car price in list and prints all matches
        If price(J) > 7501 Then
            picResults.Print Model(J); , "                  "; VeType(J); , "                  "; color(J); , "         "; mileage(J); , "          "; FormatCurrency(price(J))
        End If
    Next J
End If
End Sub

Private Sub txtType_Change()
CarType = txtType
End Sub
