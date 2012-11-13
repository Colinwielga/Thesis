VERSION 5.00
Begin VB.Form frmLicense 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   11400
      ScaleHeight     =   4455
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sample Names"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   2280
      Picture         =   "frmLicense.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "Input Desired Name to Search For"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox txtLNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Text            =   "Input Driver's License Number"
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Driver's License Background Check"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   13455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input a Name to Search For"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   11760
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   12000
      X2              =   12600
      Y1              =   3360
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   9600
      X2              =   12000
      Y1              =   3600
      Y2              =   3360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLICK TO SEARCH"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   7200
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input Desired Drivers Number to Look Up Person"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   2040
      Y1              =   6480
      Y2              =   6600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   6480
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   960
      Y1              =   6480
      Y2              =   7320
   End
   Begin VB.Image imgSearch 
      Height          =   2430
      Left            =   5640
      Picture         =   "frmLicense.frx":1C94A
      Top             =   7800
      Width           =   2145
   End
   Begin VB.Image Image1 
      Height          =   5670
      Left            =   1800
      Picture         =   "frmLicense.frx":2DAEC
      Top             =   2040
      Width           =   9630
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Licenses(1 To 100) As String, Names(1 To 100) As String, PName As String, LNumber As String, CTR As Integer
Dim Pics(1 To 100)
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009
'Objective: to offer a hypothetical program that reads a persons name or id number
'and lists if they are criminal



Private Sub Command1_Click()
'reads file into array

Dim j As Integer
'sets counter to zero

CTR = 0
    
    Open App.Path & "/Licenses.txt" For Input As #1
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Licenses(CTR), Names(CTR), Pics(CTR)
    Loop

'sets head to table

picResults2.Print "Suspected Criminals"
picResults2.Print " "
picResults2.Print "Name"; Tab(20); "License Number"
picResults2.Print "**********************************"


'lists suspects to allow user some idea of names that will produce results

j = 0
For j = 1 To CTR
picResults2.Print Names(j); Tab(20); Licenses(j)
Next j
Close #1
End Sub

Private Sub imgSearch_Click()

CTR = 0
   'reads file into arrays
    Open App.Path & "/Licenses.txt" For Input As #1
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Licenses(CTR), Names(CTR), Pics(CTR)
    Loop

Dim Found As Boolean
Dim placeCtr As Integer         'keeps track of where you are in the list
placeCtr = 0
Found = False
PName = txtName.Text
LNumber = txtLNumber.Text

'keep looking as long as you have not found what you are looking for and
' you have not reached the end of the array
Do While (Not Found) And (placeCtr < CTR)
    placeCtr = placeCtr + 1
    If (Licenses(placeCtr)) = LNumber Then

'if license is found then it jumps to the msgbox after else.

        Found = True
    ElseIf Names(placeCtr) = PName Then
'if names is found then it jumps to the msgbox after else
        Found = True
    End If
Loop

If Not Found Then

'if not then it prints that no matchs were found

    MsgBox "No criminals match this name nor license number.", , "Negative"
Else
    picResults.Picture = LoadPicture(App.Path & "\" & Pics(placeCtr))
   MsgBox Names(placeCtr) & " is a dangerous criminal. BE AWARE!!!", , "WARNING!!!"

End If
Close #1

'closes input

End Sub
'offers option to quit
Private Sub quit_Click()
End
End Sub

Private Sub return_Click()
frmLicense.Hide
frmHome.Show
End Sub

