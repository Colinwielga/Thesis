VERSION 5.00
Begin VB.Form Buildingform2 
   BackColor       =   &H00000080&
   Caption         =   "Form2"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgotomy 
      BackColor       =   &H0000FF00&
      Caption         =   "Go to my links!"
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4440
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8355
      TabIndex        =   13
      Top             =   3600
      Width           =   8415
   End
   Begin VB.CommandButton cmdidealcanoe 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click here to find your canoe builiding specifications:"
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the beginning."
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox peoplebox 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox dollarsbox 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox hoursbox 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdDUN 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by: Justin Olsen"
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "This figure too will help the program find a canoe that will best suite your situation."
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label passangers 
      BackColor       =   &H0000C000&
      Caption         =   "How many passangers will typically be riding in your canoe (1-7)?"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "This figure too will help the program find a canoe that will best suite your situation."
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "How much money do you want to spend on building your canoe?"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   $"Buildingform2.frx":0000
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "How many hours do you want to dedicate to building your canoe?"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Buildingform2.frx":0091
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Buildingform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Purpose = This form is here to ask the user a few questions about their situation and give them hints on what type of canoe would best suit them.
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Buildingform2.frm")
Private Sub cmdDUN_Click()
End
End Sub

Private Sub cmdgotomy_Click()
Buildingform2.Hide
Form2.Show
End Sub

Private Sub cmdidealcanoe_Click()
Dim hours As Single, dollars As Integer, people As Integer, A As String, B As String, C As String
picresults.Cls
hours = hoursbox.Text
If hours >= 300 Then
        A = "Working that long you can build a canoe 18 feet or longer."
    ElseIf hours >= 250 Then
        A = "Working that long you can build a canoe 16 feet or longer."
    ElseIf hours >= 225 Then
        A = "Working that long you can build a canoe 15 feet or longer."
    ElseIf hours >= 200 Then
        A = "Working that long you can build a canoe 14 feet or longer."
    ElseIf hours >= 150 Then
        A = "Working that long you can build a canoe under 12 feet long."
    Else: A = "You will need to work longer to build a canoe that you can actually ride in!"
End If
dollars = dollarsbox.Text
If dollars >= 1200 Then
        B = "You can typically build a canoe with grade A materials spending that amount of money."
    ElseIf dollars >= 1000 Then
        B = "You can typically build a canoe with grade B materials spending that amount of money."
    ElseIf dollars >= 800 Then
        B = "You can typically build a canoe with grade C materials spending that amount of money."
    ElseIf dollars >= 600 Then
        B = "Using grade D materials."
    Else: B = "If you can't spend anymore than $600 on materials, you better save some more money!"
End If
people = peoplebox.Text
Select Case people
    Case Is >= 7
        C = "With that number of people riding, you better just buy yourself a yacht!"
    Case Is >= 5
        C = "With that number of people riding you would need a canoe longer than 17 feet."
    Case Is >= 3
        C = "With that number of people riding you would need a canoe longer than 16 feet."
    Case Is = 2
        C = "With that number of people riding you would need a canoe longer than 14 feet."
    Case Is = 1
        C = "By yourself you can ride any canoe shorter than 17 feet."
    Case Else
        C = "Please enter a number of people."
End Select
picresults.Print A; Tab(1); B; Tab(1); C
End Sub

Private Sub cmdstart_Click()
Form1.Show
Buildingform2.Hide
End Sub
