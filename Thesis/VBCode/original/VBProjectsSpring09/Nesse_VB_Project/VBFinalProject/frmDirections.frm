VERSION 5.00
Begin VB.Form frmDirections 
   Caption         =   "Directions"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGet 
      Caption         =   "Show Directions"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6720
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to DAC"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   480
      ScaleHeight     =   1515
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   2640
      Width           =   8055
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Click Here to View a Map of the Capitol Complex"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox txtDirection 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter 1, 2, 3 or 4 to singify the direction you are coming from (N,S,E,W):"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Directions to the MN State Capitol:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmDirections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmDAC.Show
frmDirections.Hide
End Sub

Private Sub cmdGet_Click()
picResults.Cls

Dim Direction As Integer
Direction = txtDirection.Text

picResults.Print "The Capitol complex is to the north of I-94 just minutes from downtown St. Paul."
picResults.Print "It is accessible from the east and west on I-94, and from the north and south on I-35E."
picResults.Print
picResults.Print
picResults.Print "Your Directions:"

    If Direction = 1 Then
        picResults.Print "I-35E southbound: Exit at University Avenue. Turn right."
        picResults.Print "Go to Rice Street and turn left. Go one block, turn right and enter Parking Lot AA."
        ElseIf Direction = 2 Then
        picResults.Print "I-35E northbound: Exit at Kellogg Boulevard."
        picResults.Print "Turn left. Go to John Ireland Boulevard and turn right. Metered parking spaces line both sides of the boulevard."
        ElseIf Direction = 3 Then
        picResults.Print "From I-94 westbound: Exit at Marion Street."
        picResults.Print "Turn right. Go to Aurora Avenue and turn right. Go one block and enter Parking Lot AA."
        ElseIf Direction = 4 Then
        picResults.Print "From  I-94 eastbound: Exit at Marion Street."
        picResults.Print "Turn left. Go to Aurora Avenue and turn right. Go one block and enter Parking Lot AA."
        Else
        MsgBox "You did not enter an accurate direction. Please enter either North, South, East or West.", , "Error!"
    End If
    
End Sub

Private Sub cmdMap_Click()
frmDirections.Hide
frmMap.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

