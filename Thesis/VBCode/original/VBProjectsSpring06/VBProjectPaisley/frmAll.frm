VERSION 5.00
Begin VB.Form frmAll 
   Caption         =   "All Around Shoes"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   Picture         =   "frmAll.frx":0000
   ScaleHeight     =   594.738
   ScaleMode       =   0  'User
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      Height          =   4815
      Left            =   5640
      ScaleHeight     =   4755
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox txtRate 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2280
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
      Left            =   360
      TabIndex        =   3
      Top             =   6240
      Width           =   2775
   End
   Begin VB.CommandButton cmdRestart 
      BackColor       =   &H8000000E&
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
      Index           =   1
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
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
      Left            =   7680
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
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
      Left            =   7680
      TabIndex        =   8
      Top             =   0
      Width           =   2295
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
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   1320
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
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shoes that do it all, except Big Wall."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form is used by customers to find Sport Climbing shoes that match their individual Data.

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
    picOutput.Print "*********************************************************************"
    For pos = 1 To Size             'Searches Sorted Arrays for matches of AGR, tprice and rating with shoe style repeats to the size of the Arrays
        If Agr(pos) <> 3 Then
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
Private Sub cmdRestart_Click(Index As Integer)  'This cmdClick subroutine Shows the Main Form form and Hides the Current Form
    frmAll.Hide
    frmMain.Show
End Sub
