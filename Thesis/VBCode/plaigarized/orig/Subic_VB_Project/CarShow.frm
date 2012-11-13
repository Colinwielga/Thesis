VERSION 5.00
Begin VB.Form frmCarShow 
   BackColor       =   &H000000C0&
   Caption         =   "Car Show"
   ClientHeight    =   11940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleMode       =   0  'User
   ScaleWidth      =   11686.6
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000FF&
      Caption         =   "Hit the Gas "
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10080
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      Height          =   7455
      Left            =   1440
      Picture         =   "CarShow.frx":0000
      ScaleHeight     =   7395
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   2160
      Width           =   12015
   End
   Begin VB.Label Wellcome 
      BackColor       =   &H000000C0&
      Caption         =   "Welcome to European Car Market Show"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   11655
   End
End
Attribute VB_Name = "frmCarShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT:  VBProject(CarShow.vbp)
'AUTHOR:  Sasa Subic
'DATE:  February 25th, 2010
'PURPOSE: The purpose of this project is to provide the users with the information of latest car models produced my leading European manufufacturer.

Private Sub cmdStart_Click()
    Dim pass As Integer, tempModelS As String, tempClassS As String, tempPriceS As Single, tempMakerS As String, C As Integer, I As Integer, J As Integer
    'open the file cars.txt for input
    Open App.Path & "\cars.txt" For Input As #1
    'declare variables
    CTR = 0 'sets value of CTR to 0
    CTRS = 0 'sets value of CTRS to 0
    
'This loop reads data from a file into four arrays
'read cars.txt until End of File
Do While Not EOF(1)
    'add 1 to value of CTR to keep track of number of data lines in file
    CTR = CTR + 1
    'store data from line of text in Cars.txt as Maker(CTR), Model(CTR), Class(CTR), Price(CTR)
    Input #1, Maker(CTR), Model(CTR), Class(CTR), Price(CTR)
'return to the file to repeat the steps above
Loop
'close file
Close #1

'This loop duplicate new four arays to be used for sorting
For J = 1 To CTR
        'adds 1 to value of CTRS to keep track of number of data lines in file
        CTRS = CTRS + 1
        'declare duplicated variables
        MakerS(J) = Maker(J)
        ModelS(J) = Model(J)
        ClassS(J) = Class(J)
        PriceS(J) = Price(J)
'returns and go agin to duplicate all lines in the cars.txt file
Next J

    'sorts files according to price
    'number of passes through the list
    For pass = 1 To CTRS - 1
    'number of comparisons for each pass
    For C = 1 To CTRS - pass
        'compare adjacent prices
        If PriceS(C) > PriceS(C + 1) Then
            'swap if necessary
            tempPriceS = PriceS(C)
            PriceS(C) = PriceS(C + 1)
            PriceS(C + 1) = tempPriceS
            
            'swap if necessary
            tempMakerS = MakerS(C)
            MakerS(C) = MakerS(C + 1)
            MakerS(C + 1) = tempMakerS
            
            'swap if necessary
            tempModelS = ModelS(C)
            ModelS(C) = ModelS(C + 1)
            ModelS(C + 1) = tempModelS
            
            'swap if necessary
            tempClassS = ClassS(C)
            ClassS(C) = ClassS(C + 1)
            ClassS(C + 1) = tempClassS
        End If
    
    Next C
Next pass


    'retrieves and stores UserName in module/public
    UserName = InputBox("What's your name?", "Welcome!")
    'hides Car Show page from user
    frmCarShow.Hide
    'shows main page to user
    frmStartPage.Show
    'messagebox to user, greeting them by name and welcoming them to the Start Page
    MsgBox "Welcome to the European Car Show, " & UserName & ".", , "Greetings."
End Sub

