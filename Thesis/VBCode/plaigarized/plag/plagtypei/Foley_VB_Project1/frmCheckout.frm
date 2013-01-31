VERSION 5.00
Begin VB.Form frmCheckout
   BackColor       =   &H00004080&
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit2
      BackColor       =   &H0080FF80&
      Caption         =   "Thanks for the animals Jim, I will be back soon"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdTotal
      BackColor       =   &H0000C0C0&
      Caption         =   "Total"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd
      BackColor       =   &H0000C000&
      Caption         =   "Add another animal"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCounter
      BackColor       =   &H00FFFF80&
      Caption         =   "Return to Counter"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picResults2
      Height          =   2895
      Left            =   9000
      ScaleHeight     =   2835
      ScaleWidth      =   5355
      TabIndex        =   7
      Top             =   1680
      Width           =   5415
   End
   Begin VB.CommandButton cmdPurchase
      BackColor       =   &H0000C000&
      Caption         =   "Make a Purchase"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlphabetize
      BackColor       =   &H000000FF&
      Caption         =   "Sort by price"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrice
      BackColor       =   &H000000FF&
      Caption         =   "Sort by name"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox picResults
      Height          =   3495
      Left            =   1680
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H008080FF&
      Caption         =   "I changed my mind and I don't want an exotic animal at this time"
      Height          =   615
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrices
      Caption         =   "Prices"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1
      BackColor       =   &H00004080&
      Caption         =   "Welcome to the checkout counter"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pass, Temp As String, Position As Integer, Found As Boolean, Total As Integer, Y As Integer, subTotal As Integer
'Jimmy's Pet Store
'frmCheckout
'Jimmy Foley
'February 24, 2010
'this form is the Checkout lobby where the customer is able to see the prices, sort the animals, and make a purchase
Private Sub cmdAdd_Click() ' this allows the customer to make additional animal purchases
Dim NextAnimal As String





NextAnimal = InputBox("Enter the next animal you wish to purchase")
 I = 0
    Found = False







Do While ((Not Found) And (I < Ctr)) ' again the array is seached for a match
        I = I + 1






    If NextAnimal = Animals(I) Then
        Found = True
    End If


         ' lasdkfj dslf dslf sdkfsdlkf jds


Loop

If (Not Found) Then



     ' lasdkfj dslf dslf sdkfsdlkf jds




    MsgBox "Animal not in stock", vbCritical
Else
    MsgBox ("This animal is also a Good Choice"), vbInformation





    picResults2.Print NextAnimal; Tab(33); FormatCurrency(Prices(I))
    subTotal = subTotal + Prices(I)


         ' lasdkfj dslf dslf sdkfsdlkf jds


End If
End Sub

Private Sub cmdAlphabetize_Click() ' this bubble sort reorganizes the animals in order of prices
picResults.Cls
picResults.Print "Animal"; Tab(33); "Price"



     ' lasdkfj dslf dslf sdkfsdlkf jds


picResults.Print "****************************************************************"
    For Pass = 1 To Ctr - 1




        For Position = 1 To Ctr - Pass
            If Prices(Position) > Prices(Position + 1) Then

                Temp = Prices(Position)



                Prices(Position) = Prices(Position + 1)
                Prices(Position + 1) = Temp

                Temp = Animals(Position)





                Animals(Position) = Animals(Position + 1)
                Animals(Position + 1) = Temp

            End If
        Next Position


             ' lasdkfj dslf dslf sdkfsdlkf jds


    Next Pass

    MsgBox ("Sort Completed"), vbInformation



For Position = 1 To 8
    picResults.Print Animals(Position); Tab(33); FormatCurrency(Prices(Position))


         ' lasdkfj dslf dslf sdkfsdlkf jds



Next Position
End Sub

Private Sub cmdCounter_Click() 'this returns the user to the front of the store


     ' lasdkfj dslf dslf sdkfsdlkf jds

frmCheckout.Hide
frmEntrance.Show

End Sub

Private Sub cmdPrice_Click() ' this bubble sort reorders the animals in alphebetical order
picResults.Cls



     ' lasdkfj dslf dslf sdkfsdlkf jds


picResults.Print "Animal"; Tab(33); "Price"
picResults.Print "****************************************************************"
    For Pass = 1 To Ctr - 1

         ' lasdkfj dslf dslf sdkfsdlkf jds


        For Position = 1 To Ctr - Pass
            If Animals(Position) > Animals(Position + 1) Then


                Temp = Animals(Position)

                Animals(Position) = Animals(Position + 1)
                Animals(Position + 1) = Temp



                Temp = Prices(Position)
                Prices(Position) = Prices(Position + 1)


                Prices(Position + 1) = Temp

            End If
        Next Position
    Next Pass

    MsgBox ("Sort Completed"), vbInformation
For Position = 1 To 8

    picResults.Print Animals(Position); Tab(33); FormatCurrency(Prices(Position))
Next Position

     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub

     ' lasdkfj dslf dslf sdkfsdlkf jds




Private Sub cmdPrices_Click() ' this button inputs data from a text file into two arrays and prints the results
picResults.Print "Animal"; Tab(33); "Price"
picResults.Print "***************************************************"
Open App.Path & "\Animals2.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Animals(Ctr), Prices(Ctr)

    picResults.Print Animals(Ctr); Tab(33); FormatCurrency(Prices(Ctr))
Loop
Close #1
End Sub

Private Sub cmdPurchase_Click() 'this allows the customer to choose an animal he or she wishes to buy
Dim Purchase As String
picResults2.Cls


     ' lasdkfj dslf dslf sdkfsdlkf jds
picResults2.Print "Animal"; Tab(33); "Price"
picResults2.Print "*****************************************************"
Purchase = InputBox("Enter the Name of the Animal you wish to Purchase")


     ' lasdkfj dslf dslf sdkfsdlkf jds
    I = 0
    Found = False
Do While ((Not Found) And (I < Ctr)) ' search the array for the animal
        I = I + 1
    If Purchase = Animals(I) Then



        Found = True
    End If

         ' lasdkfj dslf dslf sdkfsdlkf jds
Loop

If (Not Found) Then ' these will let the user know if they have found a match or not
    MsgBox "Animal not in stock", vbCritical
Else
    MsgBox ("This animal is a Good Choice"), vbInformation

         ' lasdkfj dslf dslf sdkfsdlkf jds

    picResults2.Print Purchase; Tab(33); FormatCurrency(Prices(I))
    Y = Prices(I)
End If
End Sub

Private Sub cmdQuit_Click() ' I think you know what this one does by now
End

     ' lasdkfj dslf dslf sdkfsdlkf jds

End Sub

Private Sub cmdQuit2_Click()


End
End Sub

Private Sub cmdTotal_Click() ' this total botton shows the customer how much they will be spending, it adds a personal touch by using the name inputted on the entrance form



Total = subTotal + Y
picResults2.Print "******************************************************"
picResults2.Print "Congradulations "; surName; " you have made a wise investment"


picResults2.Print "Total:"; Tab(33); FormatCurrency(Total)
End Sub
