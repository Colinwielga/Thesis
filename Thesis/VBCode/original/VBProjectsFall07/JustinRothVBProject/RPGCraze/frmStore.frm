VERSION 5.00
Begin VB.Form frmStore 
   BackColor       =   &H00000000&
   Caption         =   "Store"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Picture         =   "frmStore.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Items"
      Height          =   495
      Left            =   8880
      TabIndex        =   19
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   2955
      ScaleWidth      =   5355
      TabIndex        =   16
      Top             =   3600
      Width           =   5415
      Begin VB.PictureBox picCash 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3840
         ScaleHeight     =   915
         ScaleWidth      =   1395
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton cmdSoap 
      Caption         =   "Bar Soap"
      Height          =   495
      Left            =   8880
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdJacket 
      Caption         =   "Jacket"
      Height          =   495
      Left            =   8880
      TabIndex        =   12
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdHat 
      Caption         =   "Hat"
      Height          =   495
      Left            =   8880
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackpack 
      Caption         =   "Backpack"
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPistol 
      Caption         =   "Pistol"
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGum 
      Caption         =   "Chewing Gum"
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdMatches 
      Caption         =   "Matches"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdPotion 
      Caption         =   "Magic Potion"
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   2415
      Left            =   3240
      ScaleHeight     =   2355
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlphSort 
      Caption         =   "Sort in Alphabetical Order"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdPriceSort 
      Caption         =   "Sort by Price"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Show Product Info"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Line ln2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   8760
      X2              =   10200
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line ln3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   360
      X2              =   10200
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label lblBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buy:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   735
      Left            =   8880
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   2880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblKwik 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kwik-E-Mart"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmStore
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form allows the user to view and purchase various products with their cash funds.

Option Explicit

Dim Products(1 To 15) As String, Prices(1 To 15) As Single  'Decalres the product and prices array.
Dim Total As Single

Private Sub cmdAlphSort_Click()
    Dim Pass As Integer, Pos As Integer 'Decalres the pass and pos variables.
    Dim Temp As String, Temp2 As Single 'Decalres the Temp and Temp2 variables.

    picResults.Cls  'Clears the picture box for the next click.
    
    'Sorts the displayed products into alphabetical order.
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Products(Pos) > Products(Pos + 1) Then
                Temp = Products(Pos)
                Products(Pos) = Products(Pos + 1)
                Products(Pos + 1) = Temp
            
                Temp2 = Prices(Pos)
                Prices(Pos) = Prices(Pos + 1)
                Prices(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass

    'Sets a layout up for the data.
    picResults.Print Tab(1); "Products"; Tab(30); "Prices"
    picResults.Print "*************************************************************************"

    For Pos = 1 To CTR
        picResults.Print Tab(1); Products(Pos), Tab(30); FormatCurrency(Prices(Pos))    'Prints the data into alphabetical order.

    Next Pos

End Sub

Private Sub cmdBack_Click()
    frmStore.Hide   'Goes back to the Map form.
End Sub

Private Sub cmdBackpack_Click() 'Creates a backpack product for the user to buy and adds it to the total.
    Dim Backpack As Single
        Backpack = 25
        Total = Total + Backpack
        picResults2.Print "Backpack"; Tab(20); FormatCurrency(Backpack)
End Sub

Private Sub cmdBuy_Click()  'Allows the user to purchase items with their cash funds and tells them if they are able to buy or not.
    If Total = 0 Then
        MsgBox "Please select something to purchase"
    Else
        MsgBox "Thank you for your purchase, " & N
    End If
    If Cash = 0 Then
        MsgBox "Oops! You do not have enough funds to purchase this order.", , "Insufficient funds!"
    ElseIf Total > Cash Then
        MsgBox "Oops! You do not have enough funds to purchase this order.", , "Insufficient funds!"
    End If
    If Cash > 0 Then    'Resets the cash after the purchase.
        Cash = Cash - Total
        picCash.Cls
        picCash.Print "Cash:"
        picCash.Print FormatCurrency(Cash)
    End If
End Sub

Private Sub cmdClear_Click()
    picResults2.Cls
    Total = 0
End Sub

Private Sub cmdGum_Click()  'Creates a gum product for the user to buy and adds it to the total.
    Dim Gum As Single
        Gum = 0.75
        Total = Total + Gum
        picResults2.Print "Chewing Gum"; Tab(20); FormatCurrency(Gum)
End Sub

Private Sub cmdHat_Click()  'Creates a hat product for the user to buy and adds it to the total.
    Dim Hat As Single
        Hat = 15
        Total = Total + Hat
        picResults2.Print "Hat"; Tab(20); FormatCurrency(Hat)
End Sub

Private Sub cmdJacket_Click()   'Creates a jacket product for the user to buy and adds it to the total.
    Dim Jacket As Single
        Jacket = 35
        Total = Total + Jacket
        picResults2.Print "Jacket"; Tab(20); FormatCurrency(Jacket)
End Sub

Private Sub cmdMatches_Click()  'Creates a matches product for the user to buy and adds it to the total.
    Dim Matches As Single
        Matches = 3
        Total = Total + Matches
        picResults2.Print "Matches"; Tab(20); FormatCurrency(Matches)
End Sub

Private Sub cmdPistol_Click()   'Creates a pistol product for the user to buy and adds it to the total.
    Dim Pistol As Single
        Pistol = 250
        Total = Total + Pistol
        picResults2.Print "Pistol"; Tab(20); FormatCurrency(Pistol)
End Sub

Private Sub cmdPotion_Click()   'Creates a potion product for the user to buy and adds it to the total.
    Dim Potion As Single
        Potion = 10
        Total = Total + Potion
        picResults2.Print "Magic Potion"; Tab(20); FormatCurrency(Potion)
End Sub

Private Sub cmdPriceSort_Click()    'Sorts the displayed products from greatest to least price order.
    Dim Pass As Integer, Pos As Integer
    Dim Temp As Single, Temp2 As String

    picResults.Cls  'Clears the picture box for the next output.
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Prices(Pos) < Prices(Pos + 1) Then
                Temp = Prices(Pos)
                Prices(Pos) = Prices(Pos + 1)
                Prices(Pos + 1) = Temp
            
                Temp2 = Products(Pos)
                Products(Pos) = Products(Pos + 1)
                Products(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass

    picResults.Print Tab(1); "Products"; Tab(30); "Prices"
    picResults.Print "*************************************************************************"

    For Pos = 1 To CTR
        picResults.Print Tab(1); Products(Pos), Tab(30); FormatCurrency(Prices(Pos))    'Prints the newly organized products.

    Next Pos

End Sub

Private Sub cmdRead_Click() 'Reads the product text file into the arrays and displays the user's cash.
    Dim I As Single
    
    picResults.Cls
    picCash.Cls
    
    If Cash = 0 Then
        picCash.Print "Cash:"
        picCash.Print FormatCurrency(0)
    End If
    
    If Cash > 0 Then
        picCash.Cls
        picCash.Print "Cash:"
        picCash.Print FormatCurrency(Cash)
    End If

    Open App.Path & "/products.txt" For Input As #1

    picResults.Print Tab(1); "Products"; Tab(30); "Prices"
    picResults.Print "*************************************************************************"

    CTR = 0

    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Products(CTR), Prices(CTR)
    Loop

    Close #1

    For I = 1 To CTR
        picResults.Print Tab(1); Products(I), Tab(30); FormatCurrency(Prices(I))
    Next I

End Sub

Private Sub cmdSoap_Click() 'Creates a soap product for the user to buy and adds it to the total.
    Dim Soap As Single
        Soap = 2
        Total = Total + Soap
        picResults2.Print "Bar Soap"; Tab(20); FormatCurrency(Soap)
End Sub

Private Sub cmdTotal_Click()    'Computes the total of the proucts selected.
    If Total = 0 Then
        MsgBox "Please select something to purchase"
    ElseIf Total > 0 Then
        picResults2.Print "***************************************"
        picResults2.Print "Total:", Tab(20); FormatCurrency(Total)
    End If
End Sub

Private Sub Form_Load() 'Tells the user that they don't have any cash upon entering the store.
    If Cash = 0 Then
        MsgBox "You do not have enough funds to buy anything. You can only 'window shop'."
    End If
End Sub

