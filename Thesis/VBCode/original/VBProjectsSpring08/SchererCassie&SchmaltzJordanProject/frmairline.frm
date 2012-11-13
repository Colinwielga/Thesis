VERSION 5.00
Begin VB.Form frmairline 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbusiness 
      BackColor       =   &H00FF0000&
      Caption         =   "Display Business Prices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C000C0&
      Caption         =   "Go to Next Page"
      BeginProperty Font 
         Name            =   "Charlemagne Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdweekend 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find Flights that Leave on the Weekend"
      BeginProperty Font 
         Name            =   "@BatangChe"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdsortprice 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort by Price in Ascending Order"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdalphabetical 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Alphabetically by Destination"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   3480
      ScaleHeight     =   4755
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
   End
   Begin VB.CommandButton cmddisplay 
      BackColor       =   &H000000FF&
      Caption         =   "Display Airline Flights"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lbltitiel 
      BackColor       =   &H00FFFF00&
      Caption         =   "Flight Information:"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   9255
   End
   Begin VB.Image Image2 
      Height          =   1515
      Left            =   9840
      Picture         =   "frmairline.frx":0000
      Top             =   840
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   8490
      Left            =   3240
      Picture         =   "frmairline.frx":56C1
      Top             =   2400
      Width           =   12000
   End
End
Attribute VB_Name = "frmairline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmairline
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/7/08
'Objective: Here the user can look at various flight information
'The user can push different buttons to sort our arrays, adn display them in the table alphabetically by destination, or by price in ascending order
'We also used an exhaustive search that allows the user to print out only the flights leaving on the weekend

Option Explicit

'Here we declared our golbal variables that we need for more than one sub command

Dim departure(1 To 25) As String
Dim day(1 To 25) As String
Dim destination(1 To 25) As String
Dim price(1 To 25) As Single
Dim businesscost(1 To 25) As Single

Dim tempprice As Single
Dim tempdestination As String
Dim tempdeparture As String
Dim tempday As String
Dim tempbusinesscost As Single

Private Sub cmdalphabetical_Click()

'The picture window is cleared before the new sorted list is displayed

picResults.Cls

'This code sorts the data alphabetically by the destination
'The arrays are kept parallel
'The new list is then printed

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If destination(Pos) > destination(Pos + 1) Then
            tempdestination = destination(Pos)
            destination(Pos) = destination(Pos + 1)
            destination(Pos + 1) = tempdestination
            tempprice = price(Pos)
            price(Pos) = price(Pos + 1)
            price(Pos + 1) = tempprice
            tempdeparture = departure(Pos)
            departure(Pos) = departure(Pos + 1)
            departure(Pos + 1) = tempdeparture
            tempday = day(Pos)
            day(Pos) = day(Pos + 1)
            day(Pos + 1) = tempday
            tempbusinesscost = businesscost(Pos)
            businesscost(Pos) = businesscost(Pos + 1)
            businesscost(Pos + 1) = tempbusinesscost
        End If
    Next Pos
Next Pass

picResults.Print "Departure", "Day", "Destination", "Price", "Buisness-Trip Price"
picResults.Print "****************************************************************************************************"

For J = 1 To CTR
    picResults.Print departure(J), day(J), destination(J), FormatCurrency(price(J)), FormatCurrency(businesscost(J))
Next J


End Sub


Private Sub cmdbusiness_Click()

picResults.Cls

Dim rate As Single

rate = 0.75

For J = 1 To CTR
    businesscost(J) = price(J) * rate
Next J

picResults.Print "Departure", "Day", "Destination", "Price", "Buisness-Trip Price"
picResults.Print "****************************************************************************************************"

For J = 1 To CTR
    picResults.Print departure(J), day(J), destination(J), FormatCurrency(price(J)), FormatCurrency(businesscost(J))
Next J

cmdalphabetical.Visible = True
cmdsortprice.Visible = True
cmdweekend.Visible = True

End Sub

Private Sub cmddisplay_Click()

'This code loads the airline data into four parrallel arrays

CTR = 0

Open App.Path & "\Airline.txt" For Input As #2

Do While Not EOF(2)
    CTR = CTR + 1
    Input #2, departure(CTR)
    Input #2, day(CTR)
    Input #2, destination(CTR)
    Input #2, price(CTR)
Loop

picResults.Print "Departure", "Day", "Destination", "Price"
picResults.Print "******************************************************************"

For J = 1 To CTR
    picResults.Print departure(J), day(J), destination(J), FormatCurrency(price(J))
Next J

'The user cannot load the data twice, this will prevent them from trying
'The user can sort the data only after the data is loaded into the arrays

cmddisplay.Visible = False
cmdbusiness.Visible = True

End Sub

Private Sub cmdnext_Click()

'Here the user goes to the next form

frmairline.Hide
frmCarRental.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdsortprice_Click()

'Here the picture window is cleared before the new list is printed

picResults.Cls

'This code sorts the data by price in ascending order
'The code ensures that the arrays remain parallel
'The new list is then printed

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If price(Pos) > price(Pos + 1) Then
            tempdestination = destination(Pos)
            destination(Pos) = destination(Pos + 1)
            destination(Pos + 1) = tempdestination
            tempprice = price(Pos)
            price(Pos) = price(Pos + 1)
            price(Pos + 1) = tempprice
            tempdeparture = departure(Pos)
            departure(Pos) = departure(Pos + 1)
            departure(Pos + 1) = tempdeparture
            tempday = day(Pos)
            day(Pos) = day(Pos + 1)
            day(Pos + 1) = tempday
            tempbusinesscost = businesscost(Pos)
            businesscost(Pos) = businesscost(Pos + 1)
            businesscost(Pos + 1) = tempbusinesscost
        End If
    Next Pos
Next Pass

picResults.Print "Departure", "Day", "Destination", "Price", "Business-Trip Price"
picResults.Print "******************************************************************************************************"

For J = 1 To CTR
    picResults.Print departure(J), day(J), destination(J), FormatCurrency(price(J)), FormatCurrency(businesscost(J))
Next J

End Sub

Private Sub cmdweekend_Click()

'Here we did an exhaustive search to find the flights that departed on the weekend
'We then displayed those flights

Dim found As Boolean
found = False

picResults.Cls

picResults.Print "Departure", "Day", "Destination", "Price", "Business-Trip Price"
        picResults.Print "*****************************************************************************************************************"

For J = 1 To CTR
    If Left(day(J), 1) = "S" Then
        picResults.Print departure(J), day(J), destination(J), FormatCurrency(price(J)), FormatCurrency(businesscost(J))
        found = True
    End If
Next J

If Not found Then MsgBox ("Sorry there are no flights leaving on the weekend.")

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
