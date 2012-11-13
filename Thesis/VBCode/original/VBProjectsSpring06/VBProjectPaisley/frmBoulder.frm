VERSION 5.00
Begin VB.Form frmBoulder 
   Caption         =   "Bouldering Shoe Guide"
   ClientHeight    =   8610
   ClientLeft      =   3210
   ClientTop       =   870
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   Picture         =   "frmBoulder.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   7200
   Begin VB.PictureBox picOutput 
      Height          =   2535
      Left            =   2400
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Start New Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox txtRate 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "FIND MY SHOES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblDesigner 
      BackStyle       =   0  'Transparent
      Caption         =   "Designer: Mike Paisley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblExplainYDS 
      BackStyle       =   0  'Transparent
      Caption         =   "***Enter YDS Scale to second Decimal Place. (5.05 instead of 5.5 and 5.10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblRating 
      BackStyle       =   0  'Transparent
      Caption         =   "Input the Ratings you climb in YDS Scale."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Bouldering Shoes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmBoulder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form is used by customers to find Bouldering shoes that match their individual Data.

Private Sub cmdExit_Click() 'This cmdClick subroutine Exits the Program
    End
End Sub

Private Sub cmdFind_Click()     'Finds Shoe matches based on info that will be provided by user
    Dim tprice, YDS As Single   'Itializes inputs of user
    Dim pos As Integer
    found = False               'sets found to false in case of no matches
    tprice = InputBox("Input Max Price", "Price Input") 'Input box for max price customer is willing to pay
    YDS = txtRate.Text  'Text box for input of YDS Variable by User for search
    picOutput.Cls       'Clears picoutput for next search
    picOutput.Print "Shoe Model", , "Price", "Availablity" 'Labels columns of search results
    picOutput.Print "**********************************************************************************8"
    For pos = 1 To Size         'Searches Sorted Arrays for matches of AGR, tprice and rating with shoe style repeats to the size of the Arrays
        If Agr(pos) = 3 Then
            If tprice > Price(pos) Then
                If YDS > Rating(pos) Then
                    found = True        'For verification of a search result in last step
                    picOutput.Print Names(pos), FormatCurrency(Price(pos)), Here(pos) 'Displays a matched shoe and its info
                End If
            End If
        End If
    Next pos
    If found = False Then       'Displays a MSGBOX to say no match if there is none.  this is determined by the Found Variable
        MsgBox "No Matches in inventory", , "Refine Your Search"
    End If
End Sub

Private Sub cmdRestart_Click()  'This cmdClick subroutine Shows the Main Form form and Hides the Current Form
    frmBoulder.Hide
    frmMain.Show
End Sub

