VERSION 5.00
Begin VB.Form frmSeat 
   BackColor       =   &H000000FF&
   Caption         =   "Seat"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   Picture         =   "frmSeat.frx":0000
   ScaleHeight     =   10170
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   3600
      ScaleHeight     =   3315
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   6000
      Width           =   6135
   End
   Begin VB.CommandButton cmdModels 
      BackColor       =   &H000000FF&
      Caption         =   "Check Models"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   3015
   End
   Begin VB.CommandButton cmdByPrice 
      BackColor       =   &H000000FF&
      Caption         =   "Line by Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdExeo 
      Height          =   2295
      Left            =   9960
      Picture         =   "frmSeat.frx":528C4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton cmdAlhambra 
      Height          =   2415
      Left            =   9960
      Picture         =   "frmSeat.frx":56B7D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
   Begin VB.CommandButton cmdAltea 
      Height          =   2415
      Left            =   9960
      Picture         =   "frmSeat.frx":59EF2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   3855
   End
   Begin VB.CommandButton cmdToledo 
      Height          =   2295
      Left            =   9960
      Picture         =   "frmSeat.frx":5E2BA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CommandButton cmdLeon 
      Height          =   2535
      Left            =   0
      Picture         =   "frmSeat.frx":61693
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   3375
   End
   Begin VB.CommandButton cmdIbiza 
      Height          =   2175
      Left            =   0
      Picture         =   "frmSeat.frx":64CF4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAlhambra_Click()
     'give information of maker, model and class
     MsgBox " This is Seat, Model Alhambra, Minivan", , "Model"
End Sub

Private Sub cmdAltea_Click()
     'give information of maker, model and class
     MsgBox " This is Seat, Model Altea, Crossover", , "Model"
End Sub

Private Sub cmdBack_Click()
    'hide Seat page from user
    frmSeat.Hide
    'show Start page to user
    frmStartPage.Show
End Sub
'This program will search and print desired data
Private Sub CmdByPrice_Click()
    'clear the results box of any previous text
    picResults.Cls
    'label printed data
    picResults.Print Tab(10); "Seat Models"
    picResults.Print " Model", "Class", "Price"
    'separates printed data from lables
    picResults.Print "____________________________"

    'search car makers from beginint to the end of sorted arrays
    For I = 1 To CTRS
         'search for desired maker
        If MakerS(I) = "Seat" Then
            'print matching data data from sorted arrays
            picResults.Print ModelS(I); Tab(13); ClassS(I); Tab(33); (FormatCurrency(PriceS(I)))
        End If
    
    'repeat the search in next file line
    Next I
End Sub

Private Sub cmdExeo_Click()
     'give information of maker, model and class
     MsgBox " This is Seat, Model Exeo, Mid-size Luxury", , "Model"
End Sub

Private Sub cmdIbiza_Click()
    'give information of maker, model and class
    MsgBox " This is Seat, Model Ibiza, City", , "Model"
End Sub

Private Sub cmdLeon_Click()
     'give information of maker, model and class
     MsgBox " This is Seat, Model Leon, Hatchback", , "Model"
End Sub
'This program will search and print desired data
Private Sub cmdModels_Click()
    'clear the results box of any previous text
    picResults.Cls
    'label printed data
    picResults.Print Tab(10); "Seat Models"
    picResults.Print " Model", "Class", "Price"
    'separates printed data from lables
    picResults.Print "____________________________"

    'search car makers from beginint to the end of arrays
    For I = 1 To CTR
        'search for desired maker
        If Maker(I) = "Seat" Then
                'print matching data data from arrays
                picResults.Print Model(I); Tab(13); Class(I); Tab(33); (FormatCurrency(Price(I)))
        End If
        
    'repeat the search in next file line
    Next I
End Sub

Private Sub cmdToledo_Click()
     'give information of maker, model and class
     MsgBox " This is Seat, Model Toledo, Sedan", , "Model"
End Sub

