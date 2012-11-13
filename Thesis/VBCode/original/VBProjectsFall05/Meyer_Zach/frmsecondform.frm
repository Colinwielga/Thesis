VERSION 5.00
Begin VB.Form frmsecondform 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Gun Finder"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14730
   LinkTopic       =   "Form2"
   Picture         =   "frmsecondform.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H0000FFFF&
      Caption         =   "Game"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6848
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9098
      Width           =   975
   End
   Begin VB.PictureBox descBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3555
      ScaleHeight     =   315
      ScaleWidth      =   7440
      TabIndex        =   12
      Top             =   3458
      Width           =   7500
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9120
      Width           =   975
   End
   Begin VB.PictureBox picGun 
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   4178
      ScaleHeight     =   915
      ScaleWidth      =   4515
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton cmdCamo 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort by Camouflage "
      Height          =   735
      Left            =   8618
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4298
      Width           =   1215
   End
   Begin VB.CommandButton cmdMaintenance 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort by Maintenance (1-5) (Easy-Hard)"
      Height          =   975
      Left            =   7298
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4058
      Width           =   1215
   End
   Begin VB.CommandButton cmdCost 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort by Approximate Cost"
      Height          =   735
      Left            =   5978
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4298
      Width           =   1215
   End
   Begin VB.CommandButton cmdWeight 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort by Weight"
      Height          =   735
      Left            =   4658
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4298
      Width           =   1215
   End
   Begin VB.PictureBox picData 
      Height          =   1455
      Left            =   4920
      ScaleHeight     =   1395
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   5138
      Width           =   4695
   End
   Begin VB.CommandButton cmd20 
      BackColor       =   &H000000FF&
      Caption         =   "20 Guage"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9458
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2138
      Width           =   1575
   End
   Begin VB.CommandButton cmd12 
      BackColor       =   &H000000FF&
      Caption         =   "12 Guage"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6458
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2138
      Width           =   1575
   End
   Begin VB.CommandButton cmd10 
      BackColor       =   &H000000FF&
      Caption         =   "10 Guage"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3578
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2138
      Width           =   1575
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H0000FFFF&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label lblStep 
      BackColor       =   &H000000FF&
      Caption         =   "ATTENTION: First select guage, then sort information accordingly."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2205
      TabIndex        =   11
      Top             =   1425
      Width           =   10320
   End
End
Attribute VB_Name = "frmsecondform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Gun Selector (Zach Meyer's VB Project.vbp)
'Form Name : frmsecondform (frmsecondform.frm)
'Author: Zach Meyer
'Date Written: October 26, 2005
'Objective: This form provides the user with information about
                 'different kinds of guns, and then giving the user
                 'options of how they want the information to be organized.
                 'This is basically the main page of the program.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Option Explicit
Dim a(1 To 3) As String, b(1 To 3) As Double, c(1 To 3) As Double, d(1 To 3) As Double, e(1 To 3) As String
Dim i As Integer
Dim desc As String
Dim tempstr As String
Dim tempdbl As Double

'This button shows the information of the 10 guage in the picbox
'and also shows a picture of a gun in the other picbox.
'It also shows the user what most hunters use this kind of gun for.

Private Sub cmd10_Click()
    picData.Cls
    picGun.Visible = True
    picGun.Picture = LoadPicture(App.Path & "\Images\goldcamoedit.jpg")
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    Open App.Path & "\10guage.txt" For Input As #1
        For i = 1 To 3
            Input #1, a(i), b(i), c(i), d(i), e(i)
        picData.Print a(i), b(i), c(i), d(i), e(i)
        Next i
    Input #1, desc
    descBox.Cls
    descBox.Print "This style/caliber firearm would be used for " + desc
    Close #1
End Sub

'This button shows the information of the 12 guage in the picbox
'and also shows a picture of a gun in the other picbox.
'It also shows the user what most hunters use this kind of gun for.

Private Sub cmd12_Click()
    picData.Cls
    picGun.Visible = True
    picGun.Picture = LoadPicture(App.Path & "\Images\benelli_blackeagle1crop.jpg")
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    Open App.Path & "\12guage.txt" For Input As #2
        For i = 1 To 3
            Input #2, a(i), b(i), c(i), d(i), e(i)
        picData.Print a(i), b(i), c(i), d(i), e(i)
        Next i
    Input #2, desc
    descBox.Cls
    descBox.Print "This style/caliber firearm would be used for " + desc
    Close #2
End Sub

'This button shows the information of the 20 guage in the picbox
'and also shows a picture of a gun in the other picbox.
'It also shows the user what most hunters use this kind of gun for.

Private Sub cmd20_Click()
    picData.Cls
    picGun.Visible = True
    picGun.Picture = LoadPicture(App.Path & "\Images\benelli_sport1crop.jpg")
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    Open App.Path & "\20guage.txt" For Input As #3
        For i = 1 To 3
            Input #3, a(i), b(i), c(i), d(i), e(i)
        picData.Print a(i), b(i), c(i), d(i), e(i)
        Next i
    Input #3, desc
    descBox.Cls
    descBox.Print "This style/caliber firearm would be used for " + desc
    Close #3
End Sub

'This button sorts the already selected information in the picbox.

Private Sub cmdCamo_Click()
    Dim j As Integer
    Dim i As Integer
    For j = 1 To 2
        For i = 1 To 2
            If e(i) > e(i + 1) Then
                tempstr = a(i)
                a(i) = a(i + 1)
                a(i + 1) = tempstr
                tempdbl = b(i)
                b(i) = b(i + 1)
                b(i + 1) = tempdbl
                tempdbl = c(i)
                c(i) = c(i + 1)
                c(i + 1) = tempdbl
                tempdbl = d(i)
                d(i) = d(i + 1)
                d(i + 1) = tempdbl
                tempstr = e(i)
                e(i) = e(i + 1)
                e(i + 1) = tempstr
            End If
        Next i
    Next j
    picData.Cls
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    For i = 1 To 3
        picData.Print a(i), b(i), c(i), d(i), e(i)
    Next i
End Sub

'This button sorts the already selected information in the picbox.

Private Sub cmdCost_Click()
    Dim j As Integer
    Dim i As Integer
    For j = 1 To 2
        For i = 1 To 2
            If c(i) > c(i + 1) Then
                tempstr = a(i)
                a(i) = a(i + 1)
                a(i + 1) = tempstr
                tempdbl = b(i)
                b(i) = b(i + 1)
                b(i + 1) = tempdbl
                tempdbl = c(i)
                c(i) = c(i + 1)
                c(i + 1) = tempdbl
                tempdbl = d(i)
                d(i) = d(i + 1)
                d(i + 1) = tempdbl
                tempstr = e(i)
                e(i) = e(i + 1)
                e(i + 1) = tempstr
            End If
        Next i
    Next j
    picData.Cls
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    For i = 1 To 3
        picData.Print a(i), b(i), c(i), d(i), e(i)
    Next i
End Sub

'This button will end the program.

Private Sub cmdExit_Click()
    End
End Sub

'This button will bring the user to the game form.
'It will also make all of the birds on the form appear.

Private Sub cmdGame_Click()
    frmfirstform.Hide
    frmsecondform.Hide
    frmthirdform.Show
End Sub

'This button will bring the user back to the main menu form.

Private Sub cmdMain_Click()
    frmfirstform.Show
    frmsecondform.Hide
    frmthirdform.Hide
End Sub


'This button sorts the already selected information in the picbox.

Private Sub cmdMaintenance_Click()
    Dim j As Integer
    Dim i As Integer
    For j = 1 To 2
        For i = 1 To 2
            If d(i) > d(i + 1) Then
                tempstr = a(i)
                a(i) = a(i + 1)
                a(i + 1) = tempstr
                tempdbl = b(i)
                b(i) = b(i + 1)
                b(i + 1) = tempdbl
                tempdbl = c(i)
                c(i) = c(i + 1)
                c(i + 1) = tempdbl
                tempdbl = d(i)
                d(i) = d(i + 1)
                d(i + 1) = tempdbl
                tempstr = e(i)
                e(i) = e(i + 1)
                e(i + 1) = tempstr
            End If
        Next i
    Next j
    picData.Cls
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    For i = 1 To 3
        picData.Print a(i), b(i), c(i), d(i), e(i)
    Next i
End Sub

'This button sorts the already selected information in the picbox.

Private Sub cmdWeight_Click()
    Dim j As Integer
    Dim i As Integer
    For j = 1 To 2
        For i = 1 To 2
            If b(i) > b(i + 1) Then
                tempstr = a(i)
                a(i) = a(i + 1)
                a(i + 1) = tempstr
                tempdbl = b(i)
                b(i) = b(i + 1)
                b(i + 1) = tempdbl
                tempdbl = c(i)
                c(i) = c(i + 1)
                c(i + 1) = tempdbl
                tempdbl = d(i)
                d(i) = d(i + 1)
                d(i + 1) = tempdbl
                tempstr = e(i)
                e(i) = e(i + 1)
                e(i + 1) = tempstr
            End If
        Next i
    Next j
    picData.Cls
    picData.Print "Manufacturer   Weight(Lbs.)   Price($)    Maintenance    Camouflage"
    picData.Print "***************************************************************************"
    For i = 1 To 3
        picData.Print a(i), b(i), c(i), d(i), e(i)
    Next i
End Sub

