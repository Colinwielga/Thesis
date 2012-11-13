VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   Caption         =   "The Record -- Display ad -- Options"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Grand 
      Caption         =   "Click here to see the grand total."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "options.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "options.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "options.frx":1CBE4
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Quantity 
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1155
      ScaleWidth      =   9075
      TabIndex        =   3
      Top             =   3960
      Width           =   9135
   End
   Begin VB.CommandButton Build 
      Caption         =   "Click here to have your ad built by a staff member."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Discount 
      BackColor       =   &H000080FF&
      Caption         =   "Click here to receive a package discount if you placed three or more ads."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How many weeks do you want your ad to run?"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thank you for choosing an ad with the Record. Please continue the process of placing an ad with the final options below."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 2(options.frm), Kristen Nowak, 3-14-04, The purpose of this form is to allow the user to select from several options and to determine the final cost of the ad.

Private Sub Discount_Click()
Select Case Quantity.Text
Case Is >= 7
    Total = (Total * 0.85) 'calculate a 15% discount
    Results.Print "With the packaged discount, your ad will cost "; FormatCurrency(Total); " to run it for "; Quantity.Text; " weeks."
Case 5 To 6
    Total = (Total * 0.9) 'calculate a 10% discount
    Results.Print "With the packaged discount, your ad will cost "; FormatCurrency(Total); " to run it for "; Quantity.Text; " weeks."
Case 3 To 4
    Total = (Total * 0.95) 'calculate a 5% discount
    Results.Print "With the packaged discount, your ad will cost "; FormatCurrency(Total); " to run it for "; Quantity.Text; " weeks."
End Select
Discount.Enabled = False
End Sub
Private Sub Build_Click()

Quantity.Text = Quantity

If Quantity.Text < 3 Then
    Total = ((Total / Quantity) * 1.1) * Quantity  'ad a 10% fee for building an ad
    Results.Print "The cost to have your ad built and placed in The Record is "; FormatCurrency(Total)
ElseIf Quantity.Text >= 3 Then
    Total = ((Total / Quantity) * 1.05) * Quantity 'ad a 5% fee for building an ad
    Results.Print "The cost to have your ad built and placed in The Record is "; FormatCurrency(Total)
End If
Build.Enabled = False
End Sub
Private Sub Form_Load()
Grand.Enabled = False 'disable the grand total button so that user enters the amount of weeks first
Build.Enabled = False 'disable the build ad button so that user enters the amount of weeks first
Discount.Enabled = False 'disable the discount button so that user enters the amount of weeks first
End Sub
Private Sub Grand_Click()
Results.Print
Results.Print "The grand total of your ad is "; FormatCurrency(Total); ". You have now completed the program. Thank you for your business." 'print the grand total
Grand.Enabled = False 'disable buttons
Build.Enabled = False
Discount.Enabled = False
End Sub
Private Sub Quantity_Change()
Total = Quantity.Text * Total
Results.Print "It will cost "; FormatCurrency(Total); " to place your ad for "; Quantity.Text; " weeks." 'multiply the total cost of one ad by the number of weeks
Grand.Enabled = True 'now that user has entered a quantity of ads, enable other buttons
Build.Enabled = True
If Quantity.Text >= 3 Then
    Discount.Enabled = True
End If
End Sub

Private Sub Quitbutton_Click()
End 'quit
End Sub
