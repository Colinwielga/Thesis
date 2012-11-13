VERSION 5.00
Begin VB.Form frmprincessstart 
   BackColor       =   &H00FF00FF&
   Caption         =   "PrincessBegin"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   23
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Cmdcheckout 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   22
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdIwant9 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton CmdIwant8 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdIwant7 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdIwant6 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton CmdIwant5 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdIwant4 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Cmdtotal 
      Caption         =   "Total"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   14
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton CmdIwant3 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton CmdIwant2 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdIWant1 
      Caption         =   "I Want!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   6600
      ScaleHeight     =   2595
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.PictureBox PicPrice3 
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox PicPrice2 
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox PicPrice1 
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox PicResults3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5160
      ScaleHeight     =   2475
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox PicResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3720
      ScaleHeight     =   2475
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdaccessories 
      Caption         =   "Accessories"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmddress 
      Caption         =   "Dresses"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Cmdshoes 
      Caption         =   "Shoes"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2280
      ScaleHeight     =   2475
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick out one pair of shoes, one dress, and one accessory for the princess to wear"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   21
      Top             =   1440
      Width           =   8775
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmprincessstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, The user has to pick one pair of shoes, one dress, and one purse
'Then the computer will find the subtotal, tax, and total.
'The princess will think its too little, just right, or the king will get mad if it is too much
Dim Runningtotal As Long, total As Long



Private Sub cmdaccessories_Click()
Cmdshoes.Enabled = False
cmddress.Enabled = False
cmdaccessories.Enabled = False
Cmdcheckout.Enabled = False
Cmdtotal.Enabled = True
    
    CmdIWant1.Visible = False
    CmdIwant2.Visible = False
    CmdIwant3.Visible = False
    CmdIwant4.Visible = False
    CmdIwant5.Visible = False
    CmdIwant6.Visible = False
    CmdIwant7.Visible = True
    CmdIwant8.Visible = True
    CmdIwant9.Visible = True
    
picResults1.Cls
PicResults2.Cls
PicResults3.Cls
PicPrice1.Cls
PicPrice2.Cls
PicPrice3.Cls

picResults1.Picture = LoadPicture(App.Path & "\purse1.jpg")
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
PicPrice1.Print "----------------------------------------"
PicPrice1.Print Tab(5); "$450.00"

PicResults2.Picture = LoadPicture(App.Path & "\purse2.jpg")
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicPrice2.Print "----------------------------------------"
PicPrice2.Print Tab(5); "$5.00"

PicResults3.Picture = LoadPicture(App.Path & "\purse3.jpg")
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicPrice3.Print "----------------------------------------"
PicPrice3.Print Tab(5); "$80.00"

End Sub

Private Sub Cmdcheckout_Click()
    If total <= 74.9 Then
        MsgBox "The princess is unhappy because she did not spend enough money! You are warned for next time, but be careful!", , "Alert"
    ElseIf total > 74.9 And total < 1615 Then
        MsgBox "The princess is happy with her purchases, continue onto the next leg of the journey", , "Continue"
    ElseIf total >= 1615 Then
        MsgBox "The king is upset with the princess for spending so much money! Be more careful next time!", , "Alert!"
    End If
    frmprincessstart.Hide
    FrmPrincess2.Show
    PicResults3.Cls
    PicResults2.Cls
    cmddress.Enabled = False
    cmdaccessories.Enabled = False
    Cmdshoes.Enabled = True
    picResults1.Cls
    PicPrice1.Cls
    PicPrice2.Cls
    PicPrice3.Cls
    Cmdtotal.Enabled = False
    PicResults.Cls
    CmdIWant1.Visible = True
    CmdIwant2.Visible = True
    CmdIwant3.Visible = True
    CmdIwant7.Visible = False
    CmdIwant8.Visible = False
    CmdIwant9.Visible = False
End Sub

Private Sub cmddress_Click()

Cmdshoes.Enabled = False
cmddress.Enabled = False
cmdaccessories.Enabled = True
Cmdcheckout.Enabled = False
Cmdtotal.Enabled = False


    CmdIWant1.Visible = False
    CmdIwant2.Visible = False
    CmdIwant3.Visible = False
    CmdIwant4.Visible = True
    CmdIwant5.Visible = True
    CmdIwant6.Visible = True
     CmdIwant7.Visible = False
    CmdIwant8.Visible = False
    CmdIwant9.Visible = False
picResults1.Cls
PicResults2.Cls
PicResults3.Cls
PicPrice1.Cls
PicPrice2.Cls
PicPrice3.Cls

picResults1.Picture = LoadPicture(App.Path & "\dress1.jpg")
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
PicPrice1.Print "----------------------------------------"
PicPrice1.Print Tab(5); "$100.00"

PicResults2.Picture = LoadPicture(App.Path & "\dress2.jpg")
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicPrice2.Print "----------------------------------------"
PicPrice2.Print Tab(5); "$50.00"

PicResults3.Picture = LoadPicture(App.Path & "\dress3.jpg")
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicPrice3.Print "----------------------------------------"
PicPrice3.Print Tab(5); "$1000.00"



End Sub

Private Sub CmdIWant1_Click()
    PicResults.Print "Shoes", "$15.00"
    Runningtotal = Runningtotal + 15
    
    CmdIWant1.Visible = True
    CmdIwant2.Visible = False
    CmdIwant3.Visible = False
    
    
End Sub

Private Sub CmdIwant2_Click()
    PicResults.Print "Shoes", "$250.00"
    Runningtotal = Runningtotal + 250
    
    CmdIwant2.Visible = True
    CmdIWant1.Visible = False
    CmdIwant3.Visible = False
    
End Sub

Private Sub CmdIwant3_Click()
    PicResults.Print "Shoes", "$60.00"
    Runningtotal = Runningtotal + 60
    
     CmdIwant2.Visible = False
    CmdIWant1.Visible = False
    CmdIwant3.Visible = True
End Sub

Private Sub CmdIwant4_Click()

     PicResults.Print "Dress", "$100.00"
    Runningtotal = Runningtotal + 100
    
    CmdIwant4.Visible = True
    CmdIwant5.Visible = False
    CmdIwant6.Visible = False

End Sub

Private Sub CmdIwant5_Click()
    
     PicResults.Print "Dress", "$50.00"
    Runningtotal = Runningtotal + 50
    
    CmdIwant4.Visible = False
    CmdIwant5.Visible = True
    CmdIwant6.Visible = False
End Sub

Private Sub CmdIwant6_Click()
    
     PicResults.Print "Dress", "$1000.00"
    Runningtotal = Runningtotal + 1000
    
    CmdIwant4.Visible = False
    CmdIwant5.Visible = False
    CmdIwant6.Visible = True
End Sub

Private Sub CmdIwant7_Click()
     PicResults.Print "Purse", "$450.00"
    Runningtotal = Runningtotal + 450
    
    CmdIwant7.Visible = True
    CmdIwant8.Visible = False
    CmdIwant9.Visible = False

End Sub

Private Sub CmdIwant8_Click()
    PicResults.Print "Purse", "$5.00"
    Runningtotal = Runningtotal + 5
    
    CmdIwant7.Visible = False
    CmdIwant8.Visible = True
    CmdIwant9.Visible = False
End Sub

Private Sub CmdIwant9_Click()
    PicResults.Print "Purse", "$80.00"
    Runningtotal = Runningtotal + 80
    
    CmdIwant7.Visible = False
    CmdIwant8.Visible = False
    CmdIwant9.Visible = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Cmdshoes_Click()
Cmdshoes.Enabled = False
cmddress.Enabled = True
cmdaccessories.Enabled = False
Cmdcheckout.Enabled = False
Cmdtotal.Enabled = False

    CmdIWant1.Visible = True
    CmdIwant2.Visible = True
    CmdIwant3.Visible = True
    CmdIwant4.Visible = False
    CmdIwant5.Visible = False
    CmdIwant6.Visible = False
     CmdIwant7.Visible = False
    CmdIwant8.Visible = False
    CmdIwant9.Visible = False
    
picResults1.Cls
PicResults2.Cls
PicResults3.Cls
PicPrice1.Cls
PicPrice2.Cls
PicPrice3.Cls


picResults1.Picture = LoadPicture(App.Path & "\Shoes1.jpg")
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
picResults1.Print
PicPrice1.Print "----------------------------------------"
PicPrice1.Print Tab(5); "$15.00"

PicResults2.Picture = LoadPicture(App.Path & "\Shoes2.jpg")
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicResults2.Print
PicPrice2.Print "----------------------------------------"
PicPrice2.Print Tab(5); "$250.00"


PicResults3.Picture = LoadPicture(App.Path & "\Shoes3.jpg")
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicResults3.Print
PicPrice3.Print "----------------------------------------"
PicPrice3.Print Tab(5); "$60.00"
    
End Sub


Private Sub Cmdtotal_Click()

Cmdshoes.Enabled = False
cmddress.Enabled = False
cmdaccessories.Enabled = False
Cmdcheckout.Enabled = True

Cmdtotal.Enabled = False

    PicResults.Print "--------------------------------"
PicResults.Print "Sub Total", FormatCurrency(Runningtotal)
Dim Tax As Single

Tax = Runningtotal * 0.07
Runningtotal = Tax + Runningtotal
total = Runningtotal + Tax
PicResults.Print "Tax", FormatCurrency(Tax)
PicResults.Print "Total", FormatCurrency(total)

End Sub
