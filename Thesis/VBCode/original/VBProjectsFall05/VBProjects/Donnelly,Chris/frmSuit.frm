VERSION 5.00
Begin VB.Form frmWinter 
   Caption         =   "Winter Elements"
   ClientHeight    =   8805
   ClientLeft      =   2385
   ClientTop       =   1395
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   Picture         =   "frmSuit.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   10125
   Begin VB.CommandButton cmdIce 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ice Thickness"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click to find out if how thick ice should be"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.PictureBox picIce 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   5625
      TabIndex        =   8
      Top             =   4200
      Width           =   5685
   End
   Begin VB.PictureBox picWindChill 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   7080
      Width           =   6615
   End
   Begin VB.CommandButton cmdWindChill 
      BackColor       =   &H00FFFF80&
      Caption         =   "Calculate the Windchill"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox picTemp 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   6360
      Width           =   3975
   End
   Begin VB.TextBox txtTemp 
      Height          =   735
      Left            =   8160
      TabIndex        =   2
      ToolTipText     =   "Type degrees in Farenheit"
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdTemp 
      BackColor       =   &H00FFFF80&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Go back to previous screen"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chris Donnelly"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Image imgCold 
      Height          =   3015
      Left            =   6360
      Picture         =   "frmSuit.frx":3B987
      Top             =   1080
      Width           =   3750
   End
   Begin VB.Image imgIce 
      Height          =   2160
      Left            =   360
      Picture         =   "frmSuit.frx":421F0
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Learn about the elements by clicking on the pictures below"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Convert Celcius to Fahrenheit"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
End
Attribute VB_Name = "frmWinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form has various functions that will inform the user about the winter elements
Option Explicit

'this will take the user back to the main form
Private Sub cmdBack_Click()
frmWinter.Hide
frmMain.Show
End Sub
'this subroutine will set up and search arrays to determine if ice
'thickness(user entered) is safe
Private Sub cmdIce_Click()
picIce.Cls 'clears display
Dim A, I As Integer
Dim iceArray(1 To 6) As Integer
Dim resultArray(1 To 6) As String
Dim NotFound As Boolean
Open App.Path & "\ice.txt" For Input As #1
    Do Until EOF(1)         'loads and sets up arrays
        I = I + 1
        Input #1, iceArray(I), resultArray(I)
    Loop
    Close #1
    
    A = InputBox("Enter ice thickness in inches", "Ice Thickness")
    NotFound = True 'sets boleen varible to notfound
    I = 0
        Do While (NotFound And I <= 6) 'will search through array until it finds a match, or through all 6 possible matches
        I = I + 1
            If A <= iceArray(I) Then
                NotFound = False
            End If
    Loop
        picIce.Print "Ice:"; A; " inches", ; resultArray(I)
End Sub

'this will calculate the windchill
Private Sub cmdWindChill_Click()
picWindChill.Cls 'clears display
Dim I, J, Windchill As Integer 'will classify temperature, wind speed, and the resulting windchill as integers
        I = InputBox("Enter temperature in Farenheit", "Determine Windchill")
        J = InputBox("Enter wind speed in MPH", "Determine Windchill")
        Windchill = (35.74) + (0.6215 * I) - (35.75 * (J ^ 0.16)) + (0.4275 * I * (J ^ 0.16)) 'formula for finding the windchill
        picWindChill.Print "The windchill is:"; Windchill; "degrees Farenheit" 'prints the windchill
End Sub
'this will convert to Celcius
Private Sub cmdTemp_Click()
Dim F As Single, C As Single, Temp As Single 'will classify the numbers used as single
    picTemp.Cls 'clears display
    C = txtTemp
    F = C * (9 / 5) + 32
    picTemp.Print C; " degrees Celius = "; F; "degrees Fahrenheit"
End Sub
'click on image for a message box
Private Sub imgIce_Click()
    MsgBox "Lakes and ponds provide excellent snowmobiling fun, but always make sure the ice is thick enough before you enter. Use the program below to determine how much varying thicknesses of ice can handle.", , "Ice Thickness"
End Sub
'click on image for a message box
Private Sub imgCold_Click()
    MsgBox "Being caught in the cold can be deadly. Always dress for the weather you are riding in. Don't forget to take in the windchill effect. Use the programs below to convert your temperatures to Fahrenheit and determine the windchill.", , "Caught in the Cold"
End Sub
