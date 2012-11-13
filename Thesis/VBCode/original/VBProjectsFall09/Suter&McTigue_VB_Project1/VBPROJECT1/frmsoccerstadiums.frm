VERSION 5.00
Begin VB.Form frmsoccerstadiums 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   3435
   ClientTop       =   3015
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmsoccerstadiums.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12015
   Begin VB.CommandButton cmdYesNo 
      Caption         =   "Click here to find out what to do!"
      Height          =   975
      Left            =   1200
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2760
      ScaleHeight     =   915
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton cmdTickets 
      Caption         =   "Get more tickets!"
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.OptionButton optno 
      Caption         =   "No"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.OptionButton optyes 
      Caption         =   "Yes"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear the recommendations"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to main page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdcapacity 
      BackColor       =   &H00000000&
      Caption         =   "Search by capacity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4800
      ScaleHeight     =   1995
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label lblTickets 
      Caption         =   "Have you purchased your tickets yet?"
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblstadiums 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2010 World Cup Venues"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmsoccerstadiums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'USAsoccer
'frmsoccerstadiums
'author Marty, Sean
'October 17
'This form is meant to show info about stadiums and size preference
'it is also meant to help those fans navigate to ticket purchases
Private Sub cmdcapacity_Click()
    'this cmd button will ask the user for the capacity they want for a stadium
    'it will then give the user the appropriate stadium name in a picture box
    'if no stadium capacity matches the desire of the user an appropriate message will appear.
    
    Dim capacity As Single, decision As String, coast As Boolean
    
    capacity = InputBox("Enter your desired capacity for a venue.", "Find the venue that suits you best")
    
    Select Case capacity
        Case Is > 94700
            MsgBox "The largest stadium is Soccer City in Johannesburg with a capacity of 94,700.  Please enter a new number.", , "Sorry"
        Case 69071 To 94700
            decision = InputBox("Do you want the venue to be in a city on the coast? (Must answer with a 'yes' or 'no')", , "")
            If decision = "yes" Then
                coast = True
                picresults.Print ""
                picresults.Print "The best stadium for you is Moses Mabhida Stadium in Durban."
            ElseIf decision = "no" Then
                coast = False
               picresults.Print ""
               picresults.Print "The best stadium for you is Soccer City in Johannesburg."
            Else
                MsgBox "Enter a number between 1 and 10", , "Error"
            End If
        Case 62568 To 69070
            picresults.Print ""
            picresults.Print "The best stadium for you is Green Point Stadium in Cape Town."
        Case 51761 To 62567
            picresults.Print ""
            picresults.Print "The best stadium for you is Coca-Cola Park in Johannesburg."
        Case 48071 To 51760
            picresults.Print ""
            picresults.Print "The best stadium for you is Loftus Versfeld Stadium in Pretoria (Tshwane."
        Case 48001 To 48070
            picresults.Print ""
            picresults.Print "The best stadium for you is the Free State Stadium in Bloemfontein (Mangaung)."
        Case 46001 To 48000
            picresults.Print ""
            picresults.Print "The best stadium for you is the Nelson Mandela Bay Stadium in Port Elizabeth."
        Case 44001 To 46000
            picresults.Print ""
            picresults.Print "The best stadium for you is Peter Mokaba Stadium in Polokwane."
        Case 42001 To 44000
            picresults.Print ""
            picresults.Print "The best stadium for you is Mbombela Stadium in Nelspruit."
        Case 0 To 42000
            picresults.Print ""
            picresults.Print "The best stadium for you is the Royal Bafokeng Stadium in Rustenburg."
        Case Else
            MsgBox "Please enter a positive number", , "Error"
    End Select

End Sub

Private Sub cmdclear_Click()
    'This cmd button will clear the picture box
    picresults.Cls
End Sub

Private Sub cmdreturn_Click()
    'This cmd button will return the user to the main page
    frmsoccerstadiums.Hide
    Form1.Show
End Sub


Private Sub cmdTickets_Click()
'this goes to the ticket section
frmsoccerstadiums.Hide
    frmticketinfo.Show
End Sub


Private Sub cmdYesNo_Click()
'this is a button which uses the option buttons to either print info or lead the user to a new form
If optyes.Value = True Then
    Picture1.Print "Can't wait to see you!"
    Picture1.Print "Want more Tickets?"
    Picture1.Print "Click More Tickets Button."
ElseIf optno.Value = True Then
    frmsoccerstadiums.Hide
    frmticketinfo.Show
End If

End Sub

Private Sub optno_Click()
'when the option button is clicked the picture box is cleared
Picture1.Cls
End Sub

Private Sub optyes_Click()
'when the option button is clicked the picture box is cleared
Picture1.Cls
End Sub
