VERSION 5.00
Begin VB.Form frmMagazine
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optWomen
      Caption         =   "Women"
      Height          =   375
      Left            =   11400
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.OptionButton optFood
      Caption         =   "Food"
      Height          =   375
      Left            =   11400
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.OptionButton optScience
      Caption         =   "Science"
      Height          =   375
      Left            =   11400
      TabIndex        =   8
      Top             =   4200
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdBooks
      Caption         =   "Move to the Books page"
      Height          =   735
      Left            =   11760
      TabIndex        =   7
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton CmdBack
      Caption         =   "Go back to the main menu"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton cmdSortPrice
      Caption         =   "Sort the magazine by price"
      Height          =   735
      Left            =   11400
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdExport
      Caption         =   "Export"
      Height          =   615
      Left            =   11400
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdLoad
      Caption         =   "Load Magazines"
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdTopic
      Caption         =   "Search by Catergory"
      Height          =   855
      Left            =   11400
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdName
      Caption         =   "Search by Name"
      Height          =   855
      Left            =   11400
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.PictureBox picResults3
      BackColor       =   &H00FFFF80&
      FillColor       =   &H0000FF00&
      BeginProperty Font
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7815
      Left            =   3960
      ScaleHeight     =   7755
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin VB.Image Image1
      Height          =   9360
      Left            =   -360
      Picture         =   "frmMagazine.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frmMagazine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

                'Declare variables

                Dim ctr As Integer

                Dim pos As Integer

                Dim MagazineName(1 To 100) As String

                Dim Price(1 To 100) As Single

                Dim MagaType(1 To 100) As String

                Dim NameOfMaga As String





                Private Sub CmdBack_Click()

                frmChoosing.Show

                frmMagazine.Hide

                End Sub



                Private Sub cmdBooks_Click()

                frmDetails.Show

                frmMagazine.Hide

                End Sub



                Private Sub cmdExport_Click()

                'Either open the new file or append to the current file

                Open App.Path & "\ExportItems.txt" For Append As #1



                Dim I As Integer

                For I = 1 To CurrentExp

                Print #1, ExportItem(I)

                Next I

                CurrentExp = 0

                MsgBox "Successfully Export"

                Close #1

                End Sub



                Private Sub cmdLoad_Click()

                'open the data files

                Open App.Path & "\Magazines.txt" For Input As #1

                'initialize variables

                'Read the information from the array and list them

                ctr = 0

                picResults3.Cls

                picResults3.Print " Magazine's Name "; Tab(25); " Price "; Tab(45); " Magazine Type "

                picResults3.Print "*************************************************************************************"

                CurrentExp = 0

                Do While Not EOF(1)

                ctr = ctr + 1

                Input #1, MagazineName(ctr), Price(ctr), MagaType(ctr)

                CurrentExp = CurrentExp + 1

                ExportItem(CurrentExp) = MagazineName(ctr) & ", " & Price(ctr)

                picResults3.Print MagazineName(ctr); Tab(25); FormatCurrency(Price(ctr)); Tab(45); MagaType(ctr)

                Loop

                Close #1



                End Sub



                Private Sub cmdName_Click()

                'Searh magazine by name

                Dim Found As Boolean

                NameOfMaga = InputBox(" Enter the name of the magazine ", " Name of Magazine ")

                pos = 0

                Found = False

                picResults3.Cls

                picResults3.Print " Magazine's Name "; Tab(25); " Price "; Tab(45); " Magazine Type "

                picResults3.Print "*************************************************************************************"

                CurrentExp = 0

                Do While ((Not Found) And (pos < ctr))

                pos = pos + 1

                If InStr(LCase(MagazineName(pos)), LCase(NameOfMaga)) <> 0 Then

                Found = True

                End If

                Loop

                If Found = False Then

                MsgBox " No match Found "

                Else

                picResults3.Print MagazineName(pos); Tab(25); Price(pos); Tab(45); MagaType(pos)

                CurrentExp = CurrentExp + 1

                ExportItem(CurrentExp) = MagazineName(pos) & ", " & FormatCurrency(Price(pos))

                End If







                End Sub







                Private Sub cmdSortPrice_Click()

                'Sort the prices of the magazines in ascending order

                Dim Pass As Integer

                Dim Temp As Single, TempMaga As String, TempType As String

                picResults3.Cls

                picResults3.Print " Magazine's Name "; Tab(25); " Price "; Tab(45); " Magazine Type "

                picResults3.Print "*************************************************************************************"

                For Pass = 1 To ctr - 1

                For pos = 1 To ctr - Pass

                If Price(pos) > Price(pos + 1) Then

                Temp = Price(pos)

                Price(pos) = Price(pos + 1)

                Price(pos + 1) = Temp



                TempMaga = MagazineName(pos)

                MagazineName(pos) = MagazineName(pos + 1)

                MagazineName(pos + 1) = TempMaga



                TempType = MagaType(pos)

                MagaType(pos) = MagaType(pos + 1)

                MagaType(pos + 1) = TempType



                End If

                Next pos

                Next Pass

                CurrentExp = 0

                For pos = 1 To ctr

                picResults3.Print MagazineName(pos); Tab(25); Price(pos); Tab(45); MagaType(pos)

                CurrentExp = CurrentExp + 1

                ExportItem(CurrentExp) = MagazineName(pos) & ", " & FormatCurrency(Price(pos))

                Next pos

                End Sub



                Private Sub cmdTopic_Click()

                'Search by catergory

                Dim Found As Boolean

                Dim TypeOfMaga As String



                TypeOfMaga = "Science"

                If (optFood.Value) Then

                TypeOfMaga = "Food"

                ElseIf (optWomen.Value) Then

                TypeOfMaga = "Women"

                End If



                pos = 0

                Found = False



                picResults3.Cls

                picResults3.Print " Magazine's Name "; Tab(25); " Price "

                picResults3.Print "*************************************************************************************"



                CurrentExp = 0

                For pos = 1 To ctr



                'Find keywords using InStr

                If InStr(MagaType(pos), TypeOfMaga) <> 0 Then

                Found = True

                CurrentExp = CurrentExp + 1

                ExportItem(CurrentExp) = MagazineName(pos) & ", " & Price(pos)



                picResults3.Print MagazineName(pos); Tab(25); FormatCurrency(Price(pos))

                End If

                Next pos

                If Found = False Then

                MsgBox (" No match found ")

                End If



                End Sub



                Private Sub cmdView_Click()



                End Sub





