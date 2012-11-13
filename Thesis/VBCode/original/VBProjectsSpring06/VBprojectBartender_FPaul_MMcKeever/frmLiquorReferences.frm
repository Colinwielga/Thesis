VERSION 5.00
Begin VB.Form frmLiquorReferences 
   Caption         =   "Liquor References"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "frmLiquorReferences.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Click Here to Learn About Different Liquors."
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblWhiskey 
      BackStyle       =   0  'Transparent
      Caption         =   "6. Whiskey"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblVodka 
      BackStyle       =   0  'Transparent
      Caption         =   "5. Vodka"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblTequila 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Tequila"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblRum 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Rum"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblGin 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Gin"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblBrandy 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Brandy"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblLiquor 
      BackStyle       =   0  'Transparent
      Caption         =   "How well do you know your different worldly liquors?"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2895
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "frmLiquorReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmLiquorReferences(Liquor References)
'By Fred Paul & Michael McKeever
'March 22,2006
'The Liquor references form defines the six different kinds of
'hard liquors via input boxes, with their corresponding number
'and outputs a response in a msg box.

'Declare X as Integer for number in input box.
Dim X As Integer

Private Sub cmdBack_Click()
'This button hides the liquor references form and returns the user to the bar form.
    frmLiquorReferences.Hide
    frmBar.Show
End Sub

Private Sub cmdStart_Click()
  'This button displays an input box asking for a number correlating to different
  'Liquors and displays an output from a txt file in a Msgbox.
    X = Val(InputBox("Enter Correlating Liquor Number to Learn About Its History.", "Liquor Input"))
    Select Case X
        Case 1
            MsgBox "Brandy is distilled from fermented fruit such as grapes, apricots, blackberries, or peaches.  Famous types of Brandy include Cognac and Armgnac.", , "BRANDY"
        Case 2
            MsgBox "Gin is distilled from grains, and the distinctive flavor comes from juniper berries along with other herbs and spices.", , "GIN"
        Case 3
            MsgBox "Rum Originally came from the West Indies in the 17th Century.  It is distilled from sugar cane, molasses, and caramel.", , "RUM"
        Case 4
            MsgBox "Tequila is made from the agave tequilana or blue agave plant.  Tequila is either gold or white in color.", , "TEQUILA"
        Case 5
            MsgBox "Vodka is distilled from potatoes or corn and wheat mash.  Vodka originated in the Baltic States, Poland, and Russia.", , "VODKA"
        Case 6
            MsgBox "Whiskey is distilled from grain mash of corn, rye, barley, or wheat.  The five different types of whiskey are Bourbon, Canadian Whiskey, Irish Whiskey, Rye Whiskey, and Scotch Whiskey.", , "WHISKEY"
        Case Is > 6
            MsgBox "Try entering a different number.", , "Try Again"
    End Select

End Sub



