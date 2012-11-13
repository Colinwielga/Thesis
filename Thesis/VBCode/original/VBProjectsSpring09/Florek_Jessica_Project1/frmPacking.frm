VERSION 5.00
Begin VB.Form frmPacking 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H0080FF80&
      Caption         =   "Display Recommended Packing List"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortImportance 
      BackColor       =   &H0080FF80&
      Caption         =   "Sort by Importance"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0080FF80&
      Caption         =   "Sort The List Alphabetically"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FF80&
      Caption         =   "Search The List"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "@BatangChe"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   240
      ScaleHeight     =   8715
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmPacking
'Jessica Florek
'Written: 3/20/09
'Objective: This form included loading an array of recommended items to pack
'and the importance of those items. The user can then search the list or sort the list
'alphabetically or by importance.



Option Explicit
Dim list(1 To 100) As String, j As Integer, k As Integer, importance(1 To 100) As String

Private Sub cmdBack_Click()
'brings you back to be able to select cities to visit or view your budget summary
frmPacking.Hide
frmMapCities.Show
End Sub

Private Sub cmdDisplay_Click()

picResults.Cls
'this simply displays the array as it is, displaying the item and importance
picResults.Print "Recommended Packing List"
picResults.Print
picResults.Print "Item"; Tab(25); "Importance"
picResults.Print "------------------------------------------------"


For k = 1 To j
    picResults.Print list(k); Tab(25); importance(k)
Next k

End Sub

Private Sub cmdSearch_Click()
Dim searchitem As String, found As Boolean, I As Integer

'match and stop search
picResults.Cls

searchitem = InputBox("Enter the name of an item.(Please capitalize each word.)")
'solicits input from user, then will use this input to search the list of items

I = 0
found = False
'searches list for users item
Do While (Not found) And (I < j)
    I = I + 1
    If searchitem = list(I) Then
        found = True
        MsgBox (searchitem & " is recommended to be packed.")
    End If
Loop
'if not found, this message is displayed
If (Not found) Then
    MsgBox (searchitem & " was not on the list of recommended items to pack.")
End If

End Sub

Private Sub cmdSort_Click()
Dim pass As Integer, pos As Integer, temp As String, m As Integer

picResults.Cls

picResults.Print "Packing List: Alphabetical"
picResults.Print "************************************"
picResults.Print

'bubble sorts the list alphabetically
For pass = 1 To j - 1
    For pos = 1 To j - pass
        If list(pos) > list(pos + 1) Then
            temp = list(pos)
            list(pos) = list(pos + 1)
            list(pos + 1) = temp
            temp = importance(pos)
            importance(pos) = importance(pos + 1)
            importance(pos + 1) = temp
        End If
    Next pos
Next pass
'print statement
For m = 1 To j
    picResults.Print list(m)
Next m

End Sub

Private Sub cmdSortImportance_Click()
Dim pass As Integer, pos As Integer, m As Integer, temp As String

picResults.Cls
'displays a disclaimer about the low importance items
MsgBox ("Many of the low importance items can be purchased or are provided in your hotels/hostels.")
picResults.Print "Packing List: Sorted by Importance of Items"
picResults.Print "-------------------------------------------------"
picResults.Print

m = 0
'bubble sorts and displays the list based on importance, largest number of stars to smallest. also makes sure to change out the item and importance in the same manner so the information remains accurate
For pass = 1 To j - 1
    For pos = 1 To j - pass
        If importance(pos) < importance(pos + 1) Then
            temp = importance(pos)
            importance(pos) = importance(pos + 1)
            importance(pos + 1) = temp
            temp = list(pos)
            list(pos) = list(pos + 1)
            list(pos + 1) = temp
            
        End If
    Next pos
Next pass
'print statment
For m = 1 To j
    picResults.Print list(m); Tab(25); importance(m)
Next m
End Sub

Private Sub Form_Load()
'as the form is opened the file is read into an array to be available for the display options the user can select from
Open App.Path & "\PackingList.txt" For Input As #1

Do While Not EOF(1)
    j = j + 1
    Input #1, list(j), importance(j)
Loop
Close #1
End Sub
