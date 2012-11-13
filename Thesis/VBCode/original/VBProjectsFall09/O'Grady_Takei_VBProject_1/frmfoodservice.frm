VERSION 5.00
Begin VB.Form frmfoodservice 
   BackColor       =   &H00000080&
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgohome 
      Caption         =   "Back to Home"
      Height          =   855
      Left            =   2880
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdfood7 
      Caption         =   "Bennie Bucks 325 Meal Plan (Apartment/Off campus students only)"
      Height          =   855
      Left            =   2400
      TabIndex        =   8
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdfood6 
      Caption         =   "Bennie Bucks 100 Meal Plan (Apartment/Off campus students only)"
      Height          =   855
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdfood5 
      Caption         =   "75 Meal Pass Plan (Apartment/Off campus students only)"
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdfood4 
      Caption         =   "50 Meal Pass Plan (Apartment /Off campus students only)"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdfood3 
      Caption         =   "25 Meal Pass Plan (Apartment/Off campus students only)"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdfood2 
      Caption         =   "STUDENTS LIVING IN APARTMENTS OR OFF-CAMPUS"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      Height          =   5055
      Left            =   5160
      ScaleHeight     =   4995
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdfood1 
      Caption         =   "STUDENTS LIVING IN RESIDENCE HALLS(Johnnies only)"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblfoodservice 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Choose the food service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmfoodservice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningtotal As Single, Total As Single, Tuition As Single
Dim HealthCenter As Single
Private Sub cmdfood1_Click()
'the program allows the user to picl one of the food service
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee - Health Center", FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 2166.5
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = True
cmdfood2.Enabled = False
cmdfood3.Enabled = False
cmdfood4.Enabled = False
cmdfood5.Enabled = False
cmdfood6.Enabled = False
cmdfood7.Enabled = False
End Sub
Private Sub cmdfood2_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee-Health Center", FormatCurrency(HealthCenter, 2)

runningtotal = runningtotal + 2108.5
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = True
cmdfood3.Enabled = False
cmdfood4.Enabled = False
cmdfood5.Enabled = False
cmdfood6.Enabled = False
cmdfood7.Enabled = False
End Sub

Private Sub cmdfood3_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee-Health Center", FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 318#
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = False
cmdfood3.Enabled = True
cmdfood4.Enabled = False
cmdfood5.Enabled = False
cmdfood6.Enabled = False
cmdfood7.Enabled = False
End Sub

Private Sub cmdfood4_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee-Health center", FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 432.5
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = False
cmdfood3.Enabled = False
cmdfood4.Enabled = True
cmdfood5.Enabled = False
cmdfood6.Enabled = False
cmdfood7.Enabled = False
End Sub

Private Sub cmdfood5_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Health center"; Tab(29); FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 520#
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = False
cmdfood3.Enabled = False
cmdfood4.Enabled = False
cmdfood5.Enabled = True
cmdfood6.Enabled = False
cmdfood7.Enabled = False
End Sub

Private Sub cmdfood6_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee-Health Center", FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 100#
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = False
cmdfood3.Enabled = False
cmdfood4.Enabled = False
cmdfood5.Enabled = False
cmdfood6.Enabled = True
cmdfood7.Enabled = False
End Sub

Private Sub cmdfood7_Click()
Tuition = 14489
HealthCenter = 125#
picresults.Print "Tuition"; Tab(29); FormatCurrency(Tuition, 2)
picresults.Print "Fee-Health center", FormatCurrency(HealthCenter, 2)
runningtotal = runningtotal + 325#
picresults.Print "Food Plan"; Tab(29); FormatCurrency(runningtotal, 2)
'to give the user his/her total cost of university for Fall 2009
Total = runningtotal + Tuition + HealthCenter
picresults.Print "****************************************************"
picresults.Print "Total cost is "; Tab(29); FormatCurrency(Total, 2)
cmdfood1.Enabled = False
cmdfood2.Enabled = False
cmdfood3.Enabled = False
cmdfood4.Enabled = False
cmdfood5.Enabled = False
cmdfood6.Enabled = False
cmdfood7.Enabled = True
End Sub

Private Sub cmdgohome_Click()
frmfoodservice.Hide
frmhome.Show
End Sub
