VERSION 5.00
Begin VB.Form TripPrice 
   BackColor       =   &H80000003&
   Caption         =   "How much money will you need?"
   ClientHeight    =   6030
   ClientLeft      =   3015
   ClientTop       =   2400
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8400
   Begin VB.CheckBox ChkPriceless 
      BackColor       =   &H80000003&
      Caption         =   "The Satisfaction of Traveling"
      Height          =   735
      Left            =   1920
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CheckBox ChkDrugs 
      BackColor       =   &H80000003&
      Caption         =   "Illegal Drugs"
      Height          =   495
      Left            =   1920
      TabIndex        =   22
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CheckBox ChkStreet 
      BackColor       =   &H80000003&
      Caption         =   "On The Street"
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox ChkHowMany 
      BackColor       =   &H80000003&
      Caption         =   "More"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Clear Shopping Cart"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      MaskColor       =   &H8000000D&
      TabIndex        =   15
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCompute 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      MaskColor       =   &H8000000D&
      TabIndex        =   14
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      MaskColor       =   &H8000000D&
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CheckBox ChkContinental 
      BackColor       =   &H80000003&
      Caption         =   "Just The Continental Breakfast"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox ChkFastFood 
      BackColor       =   &H80000003&
      Caption         =   "Fast Food"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox ChkExpensiveRestuarant 
      BackColor       =   &H80000003&
      Caption         =   "Expensive Restuarant"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox ChkPalace 
      BackColor       =   &H80000003&
      Caption         =   "Palace"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox ChkTequilla 
      BackColor       =   &H80000003&
      Caption         =   "Tequilla"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox ChkSombrero 
      BackColor       =   &H80000003&
      Caption         =   "Sombrero"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox ChkJersey 
      BackColor       =   &H80000003&
      Caption         =   "Soccer Jersey"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Chk3People 
      BackColor       =   &H80000003&
      Caption         =   "3 People"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Chk2People 
      BackColor       =   &H80000003&
      Caption         =   "2 People"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Chk1Person 
      BackColor       =   &H80000003&
      Caption         =   "1 Person"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox ChkHotel 
      BackColor       =   &H80000003&
      Caption         =   "Hotel"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox ChkHostel 
      BackColor       =   &H80000003&
      Caption         =   "Hostel"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000003&
      Height          =   4815
      Left            =   3360
      ScaleHeight     =   4755
      ScaleWidth      =   4875
      TabIndex        =   12
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   1800
      Y1              =   4560
      Y2              =   6000
   End
   Begin VB.Label LblGifts 
      BackColor       =   &H80000003&
      Caption         =   "Gifts?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label LblFood 
      BackColor       =   &H80000003&
      Caption         =   "Type of Food?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   1800
      Y1              =   3720
      Y2              =   4680
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Label LblStay 
      BackColor       =   &H80000003&
      Caption         =   "Where Will You Stay?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label LblPeople 
      BackColor       =   &H80000003&
      Caption         =   "Number of People Traveling?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "TripPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: TripPrice.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form: Gives the user an idea of what his/her expenses will be traveling
'A series of Checkboxes and calculations gives the user a dollar amount

Option Explicit
Dim GrandTotal As Single, People As Single, Amount As Single       'Dim the variables for the entire form
'is one person is checked, then a calculation occurs; adding to the grandtotal
Private Sub Chk1Person_Click()
People = 0
    If Chk1Person = 1 Then
        Chk2People.Enabled = False
        Chk3People.Enabled = False      'makes the other check boxes in a group not accessable due to their
        ChkHowMany.Enabled = False      'only being a possibility for one box to be checked, and one calculation
            People = 1                  'gives the value of one person
                GrandTotal = (SelectedCountry * People) + GrandTotal
    Else
        Chk2People.Enabled = True
        Chk3People.Enabled = True       'If Checkbox is not clicked, then the other checkboxes in the group are accessable
        ChkHowMany.Enabled = True
    End If
End Sub
'Same as one, but calculated for two people
Private Sub Chk2People_Click()
People = 0
    If Chk2People = 1 Then
        Chk1Person.Enabled = False
        Chk3People.Enabled = False
        ChkHowMany.Enabled = False
            People = 2
                GrandTotal = (SelectedCountry * People) + GrandTotal
    Else
        Chk1Person.Enabled = True
        Chk3People.Enabled = True
        ChkHowMany.Enabled = True
    End If
End Sub

Private Sub Chk3People_Click()
People = 0
    If Chk3People = 1 Then
        Chk2People.Enabled = False
        Chk1Person.Enabled = False
        ChkHowMany.Enabled = False
            People = 3
                GrandTotal = (SelectedCountry * People) + GrandTotal
    Else
        Chk2People.Enabled = True
        Chk1Person.Enabled = True
        ChkHowMany.Enabled = True
    End If
End Sub
'gives the user the check of a free breakfast, and nothing else
Private Sub ChkContinental_Click()
Dim Nuttin As Single
Nuttin = 0                              'Nuttin is really "Nothing", a value always set at zero
    If ChkContinental = 1 Then
        ChkExpensiveRestuarant.Enabled = False
        ChkContinental.Enabled = False
        picResults.Print "Its All Inclusive :)"     'Message box appears to remind the user that all inclusive means "It comes with the room"
            GrandTotal = (Nuttin * People) + GrandTotal
    Else
        ChkExpensiveRestuarant.Enabled = True
        ChkContinental.Enabled = True
    End If
End Sub
'This will give the user the the access of checking off drugs, it will just throw them in jail
Private Sub ChkDrugs_Click()
Dim Drugs As Single
    If ChkDrugs = 1 Then
         picResults = LoadPicture(App.Path & "\Alcatraz" & ".jpg")  'A picture shows in the picResults box of a jail cell
            MsgBox ("Good Luck in Prison")      'Message Box helping the user realize he is in prison, if the picture is not sufficient enough
                ChkJersey.Enabled = False
                ChkTequilla.Enabled = False
                ChkSombrero.Enabled = False
                ChkPriceless.Enabled = False
    Else
                ChkJersey.Enabled = True
                ChkTequilla.Enabled = True
                ChkSombrero.Enabled = True
                ChkPriceless.Enabled = True
    End If
            
End Sub
'a check box for an expensive meal
Private Sub ChkExpensiveRestuarant_Click()
Dim Restuarant As Single
Restuarant = 60                         'the meal costs 60 dollars per person
    If ChkExpensiveRestuarant = 1 Then  '1 is the same as saying "true"
        ChkContinental.Enabled = False
        ChkFastFood.Enabled = False
            GrandTotal = (People * Restuarant) + GrandTotal
    Else
        ChkContinental.Enabled = True
        ChkFastFood.Enabled = True
    End If
End Sub
'FastFood Checkbox
Private Sub ChkFastFood_Click()
Dim FastFood As Single
FastFood = 4
    If ChkFastFood = 1 Then
        ChkExpensiveRestuarant.Enabled = False
        ChkContinental.Enabled = False
            GrandTotal = (FastFood * People) + GrandTotal
    Else
        ChkExpensiveRestuarant.Enabled = True
        ChkContinental.Enabled = True
    End If
End Sub
'Living arrangements in the country
Private Sub ChkHostel_Click()
Dim Hostel As Single            'Dim Variable
Hostel = 15
    If ChkHostel = 1 Then
        ChkHotel.Enabled = False
        ChkStreet.Enabled = False       'Causes the other places of stay to become inaccessable
        ChkPalace.Enabled = False
            GrandTotal = (People * Hostel) + GrandTotal
    Else
        ChkHotel.Enabled = True
        ChkStreet.Enabled = True        'If Checkbox is not checked, then the other checkboxes are accessable
        ChkPalace.Enabled = True
    End If
End Sub
'More living arrangements
Private Sub ChkHotel_Click()
Dim Hotel As Single             'Dim Variable
Hotel = 40
    If ChkHotel = 1 Then
        ChkStreet.Enabled = False
        ChkPalace.Enabled = False
        ChkHostel.Enabled = False
            GrandTotal = (People * Hotel) + GrandTotal
    Else
        ChkStreet.Enabled = True
        ChkPalace.Enabled = True
        ChkHostel.Enabled = True
    End If
End Sub
'Gives the user the option of inputing the number of people attending this trip
Private Sub ChkHowMany_Click()
People = 0
    If ChkHowMany = 1 Then
        Chk2People.Enabled = False
        Chk3People.Enabled = False
        Chk1Person.Enabled = False
        People = InputBox("How Many People are Traveling?")     'the user enters a number and it fits right into the calculations
            GrandTotal = (SelectedCountry * People) + GrandTotal
    Else
        Chk2People.Enabled = True
        Chk3People.Enabled = True
        Chk1Person.Enabled = True
    End If
End Sub
'Gift checkbox
Private Sub ChkJersey_Click()
Dim Jersey As Single            'Dim Variable
Jersey = 30
    If ChkJersey = 1 Then
        ChkSombrero.Enabled = False
        ChkTequilla.Enabled = False         'Causes the other Gifts to become inaccessable
        ChkDrugs.Enabled = False
        ChkPriceless.Enabled = False
            Amount = InputBox("How Many Would You Like To Buy?")    'All gifts, except for Illegal Drugs will have
                GrandTotal = (Amount * Jersey) + GrandTotal
    Else
        ChkSombrero.Enabled = True
        ChkTequilla.Enabled = True
        ChkDrugs.Enabled = True
        ChkPriceless.Enabled = True
    End If
            
End Sub

Private Sub ChkPalace_Click()
Dim Palace As Single            'Dim Variable
Palace = 1000                   'Palace is a fixed price
    If ChkPalace = 1 Then
        ChkStreet.Enabled = False
        ChkHostel.Enabled = False       'Causes other checkboxes for places to stay to become inaccessable
        ChkHotel.Enabled = False
            GrandTotal = (People * Palace) + GrandTotal
    Else
        ChkStreet.Enabled = True
        ChkHostel.Enabled = True        'Causes the other checkboxes to become accessable
        ChkHotel.Enabled = True
    End If
End Sub
'A littl comic relief checkbox
Private Sub ChkPriceless_Click()
    If ChkPriceless = 1 Then
        MsgBox ("Priceless!")           'When checkbox is checked, the message "Priceless!" is shown
            ChkJersey.Enabled = False
            ChkTequilla.Enabled = False     'Other gift checkboxes become inaccessable when the Priceless checkbox is checked
            ChkDrugs.Enabled = False
            ChkSombrero.Enabled = False
    Else
            ChkJersey.Enabled = True
            ChkTequilla.Enabled = True
            ChkDrugs.Enabled = True         'All Gift Checkboxes become true
            ChkSombrero.Enabled = True
    End If
End Sub
'Sombrero gift checkbox
Private Sub ChkSombrero_Click()
Dim Sombrero As Single
Sombrero = 13               'is as a Fixed price
    If ChkSombrero = 1 Then
        ChkJersey.Enabled = False
        ChkTequilla.Enabled = False
        ChkDrugs.Enabled = False
        ChkPriceless.Enabled = False
            Amount = InputBox("How Many Would You Like To Buy?")    'User is given the opportunity to buy more than one
                GrandTotal = (Amount * Sombrero) + GrandTotal       'Grandtotal is calculated
    Else
        ChkJersey.Enabled = True
        ChkTequilla.Enabled = True
        ChkDrugs.Enabled = True
        ChkPriceless.Enabled = True
    End If
End Sub
'Street as a place to stay checkbox
Private Sub ChkStreet_Click()
Dim Street As Single
Street = 0                  'Free to live on the street
    If ChkStreet = 1 Then
        ChkHostel.Enabled = False
        ChkHotel.Enabled = False
        ChkPalace.Enabled = False
            GrandTotal = (People * Street) + GrandTotal
                MsgBox ("I Think You Will Get Mugged")      'Messagebox lets the user know of the dangers of sleeping on the street
    Else
        ChkHostel.Enabled = True
        ChkHotel.Enabled = True
        ChkPalace.Enabled = True
    End If
End Sub
'Tequilla as a gift checkbox
Private Sub ChkTequilla_Click()
Dim Tequilla As Single          'Dim Variable
Tequilla = 8                    'Price is fixed
    If ChkTequilla = 1 Then
        ChkJersey.Enabled = False
        ChkSombrero.Enabled = False         'Other Checkboxes become False when this checkbox is checked
        ChkDrugs.Enabled = False
        ChkPriceless.Enabled = False
            Amount = InputBox("How Many Would You Like To Buy?")
                GrandTotal = (Amount * Tequilla) + GrandTotal
    Else
        ChkJersey.Enabled = True
        ChkSombrero.Enabled = True
        ChkDrugs.Enabled = True
        ChkPriceless.Enabled = True
    End If
        
End Sub
'Brings the user back to the Travel Form
Private Sub cmdBack_Click()
Travel.Show
TripPrice.Hide
End Sub
'Clears the Grandtotal incase the user changes his mind.
Private Sub CmdClear_Click()
GrandTotal = 0
End Sub
'Computes the Grandtotal
Private Sub cmdCompute_Click()
picResults = Nothing            'Disposes of the picture in the print box
picResults.Cls
picResults.Print "Total Trip Price"
picResults.Print "**************************************"
picResults.Print FormatCurrency(GrandTotal)         'Formats the Grandtotal into a currency.
End Sub

