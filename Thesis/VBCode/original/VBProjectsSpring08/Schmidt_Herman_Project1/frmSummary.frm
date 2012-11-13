VERSION 5.00
Begin VB.Form frmSummary 
   BackColor       =   &H00FF00FF&
   Caption         =   "Trip Summary"
   ClientHeight    =   8175
   ClientLeft      =   1545
   ClientTop       =   1920
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   12000
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00FFFF00&
      Caption         =   "$$ Compute Your Trip Total $$"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   3720
      ScaleHeight     =   5115
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox txtDestination 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Summary"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label lblChooseNumber 
      BackColor       =   &H00FF00FF&
      Caption         =   "<= Enter a number (1-6)          for your              desired          destination before you click the compute button!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblColorado 
      BackStyle       =   0  'Transparent
      Caption         =   "5.  Denver, Colorado"
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label lblFlorida 
      BackStyle       =   0  'Transparent
      Caption         =   "4.  Orlando, Florida"
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblNewYork 
      BackStyle       =   0  'Transparent
      Caption         =   "3.  New York City, New York"
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lblCalifornia 
      BackStyle       =   0  'Transparent
      Caption         =   "2.  San Diego, California"
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblHawaii 
      BackStyle       =   0  'Transparent
      Caption         =   "1.  Honolulu, Hawaii"
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: Summary
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform the users of the costs that they could expect to pay on the
'           trip that they desire.

Private Sub cmdCompute_Click()
'Declare varialbes needed for the command button.
Dim destinationNumber As Integer
Dim plane As String
Dim rental As String
Dim hotel As String
Dim activity As String
Dim salestax As Single
Dim subtotal As Long
Dim total As Long
Dim hotelavg As Single
Dim rentalavg As Single
Dim activityavg As Single
Dim planeavg As Single

'Declare where the values for destinationNumber will come from.
destinationNumber = txtDestination.Text

'Makes sure that what is written in the input boxes is what is typed in the input boxes.
MsgBox ("Make sure that you type your options exactly as they are written!!!")

'Used to find where they are going, if user is going to stay in a hotel, if user is going
'to rent a car, if user is going to fly to destination, and if the user is going to
'do any activities while on vacation.  Then it will print the averages of these and
'give the user a subtotal, sales tax, and a total that is made up of averages.
    If destinationNumber = 1 Then
        plane = InputBox("Are you going to be flying to your destination: Yes or No.")
            If plane = "Yes" Then
                    planeavg = HPave
                ElseIf plane = "No" Then
                    planeavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        rental = InputBox("Choose a Car Rental Company that you would like to use: Avis, Alamo, National.")
            If rental = "Avis" Then
                    rentalavg = HAvisavg
                ElseIf rental = "Alamo" Then
                    rentalavg = HAlamoavg
                ElseIf rental = "National" Then
                    rentalavg = HNationalavg
                Else
                    MsgBox ("Please enter the Car Rental Companies as they are written!")
            End If
        hotel = InputBox("Would you like to stay in a Hotel?  Yes or No.")
            If hotel = "Yes" Then
                    hotelavg = Havg
                ElseIf hotel = "No" Then
                    hotelavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        activity = InputBox("Are you going to be doing any activities while on vacation?  Yes or No.")
            If activity = "Yes" Then
                    activityavg = Have
                ElseIf activity = "No" Then
                    activityavg = 0
                Else
                    MsgBox ("Please enter Yes of No as they are written!")
            End If
 'If not the first destination, then goes to the this one.
    ElseIf destinationNumber = 2 Then
        plane = InputBox("Are you going to be flying to your destination: Yes or No.")
            If plane = "Yes" Then
                    planeavg = SDPave
                ElseIf plane = "No" Then
                    planeavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        rental = InputBox("Choose a Car Rental Company that you would like to use: Avis, Alamo, National.")
            If rental = "Avis" Then
                    rentalavg = SDAvisavg
                ElseIf rental = "Alamo" Then
                    rentalavg = SDAlamoavg
                ElseIf rental = "National" Then
                    rentalavg = SDNationalavg
                Else
                    MsgBox ("Please enter the Car Rental Companies as they are written!")
            End If
        hotel = InputBox("Would you like to stay in a Hotel?  Yes or No.")
            If hotel = "Yes" Then
                    hotelavg = SDavg
                ElseIf hotel = "No" Then
                    hotelavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        activity = InputBox("Are you going to be doing any activities while on vacation?  Yes or No.")
            If activity = "Yes" Then
                    activityavg = SDave
                ElseIf activity = "No" Then
                    activityavg = 0
                Else
                    MsgBox ("Please enter Yes of No as they are written!")
            End If
'If not the sencond either, it goes to this one.
    ElseIf destinationNumber = 3 Then
        plane = InputBox("Are you going to be flying to your destination: Yes or No.")
            If plane = "Yes" Then
                    planeavg = NYCPave
                ElseIf plane = "No" Then
                    planeavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        rental = InputBox("Choose a Car Rental Company that you would like to use: Avis, Alamo, National.")
            If rental = "Avis" Then
                    rentalavg = NYCAvisavg
                ElseIf rental = "Alamo" Then
                    rentalavg = NYCAlamoavg
                ElseIf rental = "National" Then
                    rentalavg = NYCNationalavg
                Else
                    MsgBox ("Please enter the Car Rental Companies as they are written!")
            End If
        hotel = InputBox("Would you like to stay in a Hotel?  Yes or No.")
            If hotel = "Yes" Then
                    hotelavg = NYCavg
                ElseIf hotel = "No" Then
                    hotelavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        activity = InputBox("Are you going to be doing any activities while on vacation?  Yes or No.")
            If activity = "Yes" Then
                    activityavg = NYCave
                ElseIf activity = "No" Then
                    activityavg = 0
                Else
                    MsgBox ("Please enter Yes of No as they are written!")
            End If
'If not the first 3, then goes to this one.
    ElseIf destinationNumber = 4 Then
        plane = InputBox("Are you going to be flying to your destination: Yes or No.")
            If plane = "Yes" Then
                    planeavg = OPave
                ElseIf plane = "No" Then
                    planeavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        rental = InputBox("Choose a Car Rental Company that you would like to use: Avis, Alamo, National.")
            If rental = "Avis" Then
                    rentalavg = OAvisavg
                ElseIf rental = "Alamo" Then
                    rentalavg = OAlamoavg
                ElseIf rental = "National" Then
                    rentalavg = ONationalavg
                Else
                    MsgBox ("Please enter the Car Rental Companies as they are written!")
            End If
        hotel = InputBox("Would you like to stay in a Hotel?  Yes or No.")
            If hotel = "Yes" Then
                    hotelavg = Oavg
                ElseIf hotel = "No" Then
                    hotelavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        activity = InputBox("Are you going to be doing any activities while on vacation?  Yes or No.")
            If activity = "Yes" Then
                    activityavg = Oave
                ElseIf activity = "No" Then
                    activityavg = 0
                Else
                    MsgBox ("Please enter Yes of No as they are written!")
            End If
 'If the first 4 don't work it goes to this, the last option.
    ElseIf destinationNumber = 5 Then
        plane = InputBox("Are you going to be flying to your destination: Yes or No.")
            If plane = "Yes" Then
                    planeavg = DPave
                ElseIf plane = "No" Then
                    planeavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        rental = InputBox("Choose a Car Rental Company that you would like to use: Avis, Alamo, National.")
            If rental = "Avis" Then
                    rentalavg = DAvisavg
                ElseIf rental = "Alamo" Then
                    rentalavg = DAlamoavg
                ElseIf rental = "National" Then
                    rentalavg = DNationalavg
                Else
                    MsgBox ("Please enter the Car Rental Companies as they are written!")
            End If
        hotel = InputBox("Would you like to stay in a Hotel?  Yes or No.")
            If hotel = "Yes" Then
                    hotelavg = Davg
                ElseIf hotel = "No" Then
                    hotelavg = 0
                Else
                    MsgBox ("Please enter Yes or No as they are written!")
            End If
        activity = InputBox("Are you going to be doing any activities while on vacation?  Yes or No.")
            If activity = "Yes" Then
                    activityavg = Dave
                ElseIf activity = "No" Then
                    activityavg = 0
                Else
                    MsgBox ("Please enter Yes of No as they are written!")
            End If
    End If

'Tells user there will be a 6% sales tax.
MsgBox ("Sales tax of 6% will be added to your subtotal.")

'Prints the info that the user puts in.
picResults.Cls
picResults.Print "The average costs to go on your trip will be:"
picResults.Print "--------------------------------------------------------------------------------------------------"
picResults.Print
picResults.Print "Flight Price (on average):", FormatCurrency(planeavg)
picResults.Print "Car Rental (on average):", FormatCurrency(rentalavg)
picResults.Print "Hotel (on average):", FormatCurrency(hotelavg)
picResults.Print "Activities (on average):", FormatCurrency(activityavg)

'Calculates and prints the information about the subtotal, sales tax, and total
picResults.Print
picResults.Print "--------------------------------------------------------------------------------------------------"
    subtotal = activityavg + hotelavg + planeavg + rentalavg
picResults.Print "Subtotal:", FormatCurrency(subtotal, 2)
    salestax = subtotal * 0.06
picResults.Print "Sales Tax:", FormatCurrency(salestax, 2)
    total = subtotal + salestax
picResults.Print "Total:", FormatCurrency(total, 2)
picResults.Print
picResults.Print "Thank you for using the Travel Agency, now you may select another trip to summarize, go back to the"
picResults.Print "homepage and exit, or go back to the homepage and look at other destinations."
End Sub

Private Sub cmdMain_Click()
'When button is clicked, the Home Page hides and the Summary form hides.
frmHome.Show
frmSummary.Hide
End Sub

Private Sub Form_Load()

'Informs the user to make sure that they are done looking at their destinations before going futher.
MsgBox ("In order to do your trip summary you must first look at all of things your city has to offer!  If you haven't, then go back to the Homepage and look at all the information that your destination has to offer!")

End Sub
