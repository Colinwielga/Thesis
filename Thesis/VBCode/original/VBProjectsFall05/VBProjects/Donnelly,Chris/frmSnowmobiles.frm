VERSION 5.00
Begin VB.Form frmSnowmobiles 
   BackColor       =   &H80000013&
   Caption         =   "Snowmobiles"
   ClientHeight    =   8535
   ClientLeft      =   2385
   ClientTop       =   1395
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10845
   Begin VB.PictureBox picOutput 
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   2160
      Width           =   10575
   End
   Begin VB.CommandButton cmdMatch 
      BackColor       =   &H80000002&
      Caption         =   "Show your matching sleds"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   4935
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by Price"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by Type"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrand 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by Brand"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H80000002&
      Caption         =   "Show All Snowmobiles"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
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
      Left            =   4560
      TabIndex        =   9
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   2250
      Left            =   7920
      Picture         =   "frmSnowmobiles.frx":0000
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   0
      Picture         =   "frmSnowmobiles.frx":B6A5
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmSnowmobiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this will display and sort the snowmobiles and information I put together
'in a text file
Option Explicit
Dim Temp, Temp2, Temp3, Temp4, Temp5 As String
Dim Pass, I As Integer
Dim SledArray(1 To 19) As String
Dim BrandArray(1 To 19) As String
Dim TypeArray(1 To 19) As String
Dim PriceArray(1 To 19) As Single
Dim EngineArray(1 To 19) As Single 'dims my four columns of information as arrays

Private Sub cmdBack_Click()
'click to go back to the Ride form
frmSnowmobiles.Hide
frmRide.Show
End Sub

Private Sub cmdDisplay_Click()
'this subroutine will display all of the information as I have it in the text file
Open App.Path & "\Snowmobiles.txt" For Input As #1 'opens text file
picOutput.Print "Snowmobile", , "Brand", , "Type", , "Cost (MSRP)", , "Engine Size" 'prints out a header on each of the four columns
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    For I = 1 To 19 'there are 19 rows in text file
        Input #1, SledArray(I), BrandArray(I), TypeArray(I), PriceArray(I), EngineArray(I)
        picOutput.Print SledArray(I), , BrandArray(I), , TypeArray(I), , "$ "; PriceArray(I), , EngineArray(I)
    Next I
End Sub
'this subroutine will sort the snowmobiles based on their brand name
Private Sub cmdBrand_Click()
picOutput.Cls
picOutput.Print "Snowmobile", , "Brand", , "Type", , "Price (MSRP)", , " Engine Size"
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        For Pass = 1 To 19
            For I = 1 To 19 - Pass
                If BrandArray(I) < BrandArray(I + 1) Then
                    Temp = BrandArray(I)
                    BrandArray(I) = BrandArray(I + 1)
                    BrandArray(I + 1) = Temp
                    Temp2 = SledArray(I)
                    SledArray(I) = SledArray(I + 1)
                    SledArray(I + 1) = Temp2
                    Temp3 = TypeArray(I)
                    TypeArray(I) = TypeArray(I + 1)
                    TypeArray(I + 1) = Temp3
                    Temp4 = PriceArray(I)
                    PriceArray(I) = PriceArray(I + 1)
                    PriceArray(I + 1) = Temp4
                    Temp5 = EngineArray(I)
                    EngineArray(I) = EngineArray(I + 1)
                    EngineArray(I + 1) = Temp5
                End If
            Next I
            picOutput.Print SledArray(I), , BrandArray(I), , TypeArray(I), , "$"; PriceArray(I), , EngineArray(I)
        Next Pass
End Sub
'this subroutine will bubble sort the snowmobiles based on price
Private Sub cmdPrice_Click()
picOutput.Cls
picOutput.Print "Snowmobile", , "Brand", , "Type", , "Price (MSRP)", , "Engine Size"
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        For Pass = 1 To 19
            For I = 1 To 19 - Pass
                If PriceArray(I) < PriceArray(I + 1) Then
                    Temp = PriceArray(I)
                    PriceArray(I) = PriceArray(I + 1)
                    PriceArray(I + 1) = Temp
                    Temp2 = SledArray(I)
                    SledArray(I) = SledArray(I + 1)
                    SledArray(I + 1) = Temp2
                    Temp3 = TypeArray(I)
                    TypeArray(I) = TypeArray(I + 1)
                    TypeArray(I + 1) = Temp3
                    Temp4 = BrandArray(I)
                    BrandArray(I) = BrandArray(I + 1)
                    BrandArray(I + 1) = Temp4
                    Temp5 = EngineArray(I)
                    EngineArray(I) = EngineArray(I + 1)
                    EngineArray(I + 1) = Temp5
                End If
            Next I
            picOutput.Print SledArray(I), , BrandArray(I), , TypeArray(I), , "$"; PriceArray(I), , EngineArray(I)
        Next Pass
End Sub

'goes back to main menu
Private Sub cmdMain_Click()
frmSnowmobiles.Hide
frmMain.Show
End Sub

'this subroutine sorts the snowmobile based on their type
Private Sub cmdType_Click()
picOutput.Cls
picOutput.Print "Snowmobile", , "Brand", , "Type", , "Price (MSRP)", , "Engine Size"
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        For Pass = 1 To 19
            For I = 1 To 19 - Pass
                If TypeArray(I) < TypeArray(I + 1) Then
                    Temp = TypeArray(I)
                    TypeArray(I) = TypeArray(I + 1)
                    TypeArray(I + 1) = Temp
                    Temp2 = SledArray(I)
                    SledArray(I) = SledArray(I + 1)
                    SledArray(I + 1) = Temp2
                    Temp3 = BrandArray(I)
                    BrandArray(I) = BrandArray(I + 1)
                    BrandArray(I + 1) = Temp3
                    Temp4 = PriceArray(I)
                    PriceArray(I) = PriceArray(I + 1)
                    PriceArray(I + 1) = Temp4
                    Temp5 = EngineArray(I)
                    EngineArray(I) = EngineArray(I + 1)
                    EngineArray(I + 1) = Temp5
                End If
            Next I
            picOutput.Print SledArray(I), , BrandArray(I), , TypeArray(I), , "$"; PriceArray(I), , EngineArray(I)
        Next Pass
End Sub
'this program will display all of the snowmobiles that have the matching type
'and matching speed (engine size) that user entered on the ride form
Private Sub cmdMatch_Click()
picOutput.Cls
picOutput.Print "Snowmobile", , "Brand", , "Type", , "Cost (MSRP)", , "Engine Size" 'prints out a header on each of the four columns
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
I = 0 'sets a counter to 0
    Do Until I = 19 'will go through all 19 rows of my file
        I = I + 1
            If selectType = TypeArray(I) And (Speed = EngineArray(I) Or Speed1 = EngineArray(I) Or Speed2 = EngineArray(I)) Then 'must match user's type and speed
                picOutput.Print SledArray(I), , BrandArray(I), , TypeArray(I), , "$"; PriceArray(I), , EngineArray(I)
            End If
    Loop
picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
picOutput.Print , , "Refer to your local dealer for more snowmobiles and more information and get ready for some winter fun!"
End Sub
