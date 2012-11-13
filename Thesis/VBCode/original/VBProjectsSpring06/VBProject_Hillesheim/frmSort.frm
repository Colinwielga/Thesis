VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H80000001&
   Caption         =   "Sort Data"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortBB 
      Caption         =   "Sort by Battleships"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortCA 
      Caption         =   "Sort by Cruisers"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortDD 
      Caption         =   "Sort by Destroyers"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortCV 
      Caption         =   "Sort by Carriers"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Stats"
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.PictureBox picDisplay 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   10395
      TabIndex        =   2
      Top             =   360
      Width           =   10455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Previous Page"
      Height          =   615
      Left            =   7680
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   9240
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000001&
      Caption         =   "Click on Load Stats to view battle statisitcs. Click on ship classification to rank battles by losses of that class"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "By Jacob Hillesheim"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   7920
      Width           =   3495
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'Data Sorting (frmSort.frm)
'Jacob Hillesheim
'March 20,2006
'The purpose of this form is for the user to be able to manipulate the data and see
'which battles more of a certain type of ship was sunk in.
Option Explicit
Dim pos, pass As Integer
Dim tempACV, tempABB, tempACA, tempADD, tempJCV, tempJBB, tempJCA, tempJDD As Integer
Dim tempBattle As String

Private Sub cmdLoad_Click()
    'clears picture
    picDisplay.Cls
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'displays headings
    picDisplay.Print "HERE ARE EIGHT FAMOUS NAVAL BATTLES OF WORLD WAR II AND CORRESPONDING LOSSES."
    picDisplay.Print
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'prints information in arrays pertaining to American losses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'displays headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'displays information in arrays pertaining to Japanese losses
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub
Private Sub cmdQuit_Click()
    'ends program
    End
End Sub
Private Sub cmdReturn_Click()
    'returns user to Battle Page
    frmBattle.Show
    frmSort.Hide
End Sub
Private Sub cmdSortBB_Click()
    'clears display box
    picDisplay.Cls
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'Displays headings
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    'orders American battles and stats by most battleships lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If ABB(pos) < ABB(pos + 1) Then
                tempABB = ABB(pos)
                ABB(pos) = ABB(pos + 1)
                ABB(pos + 1) = tempABB
                tempACV = ACV(pos)
                ACV(pos) = ACV(pos + 1)
                ACV(pos + 1) = tempACV
                tempACA = ACA(pos)
                ACA(pos) = ACA(pos + 1)
                ACA(pos + 1) = tempACA
                tempADD = ADD(pos)
                ADD(pos) = ADD(pos + 1)
                ADD(pos + 1) = tempADD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
        
    'prints headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'orders Japanese battles and stats by most battleships lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If JBB(pos) < JBB(pos + 1) Then
                tempJBB = JBB(pos)
                JBB(pos) = JBB(pos + 1)
                JBB(pos + 1) = tempJBB
                tempJCV = JCV(pos)
                JCV(pos) = JCV(pos + 1)
                JCV(pos + 1) = tempJCV
                tempJCA = JCA(pos)
                JCA(pos) = JCA(pos + 1)
                JCA(pos + 1) = tempJCA
                tempJDD = JDD(pos)
                JDD(pos) = JDD(pos + 1)
                JDD(pos + 1) = tempJDD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub

Private Sub cmdSortCA_Click()
    'Clears display
    picDisplay.Cls
    
    'inputs data from file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'displays headings
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    'orders American battles and stats by most cruisers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If ACA(pos) < ACA(pos + 1) Then
                tempACA = ACA(pos)
                ACA(pos) = ACA(pos + 1)
                ACA(pos + 1) = tempACA
                tempACV = ACV(pos)
                ACV(pos) = ACV(pos + 1)
                ACV(pos + 1) = tempACV
                tempABB = ABB(pos)
                ABB(pos) = ABB(pos + 1)
                ABB(pos + 1) = tempABB
                tempADD = ADD(pos)
                ADD(pos) = ADD(pos + 1)
                ADD(pos + 1) = tempADD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'prints headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    'orders Japanese battles and stats by most cruisers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If JCA(pos) < JCA(pos + 1) Then
                tempJCA = JCA(pos)
                JCA(pos) = JCA(pos + 1)
                JCA(pos + 1) = tempJCA
                tempJCV = JCV(pos)
                JCV(pos) = JCV(pos + 1)
                JCV(pos + 1) = tempJCV
                tempJBB = JBB(pos)
                JBB(pos) = JBB(pos + 1)
                JBB(pos + 1) = tempJBB
                tempJDD = JDD(pos)
                JDD(pos) = JDD(pos + 1)
                JDD(pos + 1) = tempJDD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub
Private Sub cmdSortCV_Click()
    'clears display
    picDisplay.Cls
    
    'inputs stats from data file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'prints headings
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    'orders American battles and stats by most carriers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If ACV(pos) < ACV(pos + 1) Then
                tempACV = ACV(pos)
                ACV(pos) = ACV(pos + 1)
                ACV(pos + 1) = tempACV
                tempABB = ABB(pos)
                ABB(pos) = ABB(pos + 1)
                ABB(pos + 1) = tempABB
                tempACA = ACA(pos)
                ACA(pos) = ACA(pos + 1)
                ACA(pos + 1) = tempACA
                tempADD = ADD(pos)
                ADD(pos) = ADD(pos + 1)
                ADD(pos + 1) = tempADD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'prints headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    'orders Japanese battles and stats by most carriers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If JCV(pos) < JCV(pos + 1) Then
                tempJCV = JCV(pos)
                JCV(pos) = JCV(pos + 1)
                JCV(pos + 1) = tempJCV
                tempJBB = JBB(pos)
                JBB(pos) = JBB(pos + 1)
                JBB(pos + 1) = tempJBB
                tempJCA = JCA(pos)
                JCA(pos) = JCA(pos + 1)
                JCA(pos + 1) = tempJCA
                tempJDD = JDD(pos)
                JDD(pos) = JDD(pos + 1)
                JDD(pos + 1) = tempJDD
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub
Private Sub cmdSortDD_Click()
    'clears picture box
    picDisplay.Cls
    
    'inputs stats from data file into arrays
    Open App.Path & "\battles.txt" For Input As #1
        For pos = 1 To 8
            Input #1, battles(pos), ACV(pos), ABB(pos), ACA(pos), ADD(pos), JCV(pos), JBB(pos), JCA(pos), JDD(pos)
        Next pos
    Close #1
    
    'displays headings
    picDisplay.Print "AMERICAN LOSSES"
    picDisplay.Print "Battle "; Tab(35); "US Carriers", "    US Battleships             ", "  US Cruisers", , "  US Destroyers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    'orders American battles and stats by most destroyers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If ADD(pos) < ADD(pos + 1) Then
                tempADD = ADD(pos)
                ADD(pos) = ADD(pos + 1)
                ADD(pos + 1) = tempADD
                tempACV = ACV(pos)
                ACV(pos) = ACV(pos + 1)
                ACV(pos + 1) = tempACV
                tempABB = ABB(pos)
                ABB(pos) = ABB(pos + 1)
                ABB(pos + 1) = tempABB
                tempACA = ACA(pos)
                ACA(pos) = ACA(pos + 1)
                ACA(pos + 1) = tempACA
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); ACV(pos); Tab(65); ABB(pos); Tab(90); ACA(pos); Tab(120); ADD(pos)
    Next pos
    
    'displays headings
    picDisplay.Print
    picDisplay.Print "JAPANESE LOSSES"
    picDisplay.Print "Battle "; Tab(32); "Japanese Carriers", "Japanese Battleships", "Japanese Cruisers", "Japanese Destoryers"
    picDisplay.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    'orders Japanese battles and stats by most destroyers lost
    For pass = 1 To 7
        For pos = 1 To (8 - pass)
            If JDD(pos) < JDD(pos + 1) Then
                tempJDD = JDD(pos)
                JDD(pos) = JDD(pos + 1)
                JDD(pos + 1) = tempJDD
                tempJCV = JCV(pos)
                JCV(pos) = JCV(pos + 1)
                JCV(pos + 1) = tempJCV
                tempJBB = JBB(pos)
                JBB(pos) = JBB(pos + 1)
                JBB(pos + 1) = tempJBB
                tempJCA = JCA(pos)
                JCA(pos) = JCA(pos + 1)
                JCA(pos + 1) = tempJCA
                tempBattle = battles(pos)
                battles(pos) = battles(pos + 1)
                battles(pos + 1) = tempBattle
            End If
        Next pos
    Next pass
    
    'prints newly ordered arrays
    For pos = 1 To 8
        picDisplay.Print battles(pos); Tab(40); JCV(pos); Tab(65); JBB(pos); Tab(90); JCA(pos); Tab(120); JDD(pos)
    Next pos
End Sub

