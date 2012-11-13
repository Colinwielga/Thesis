VERSION 5.00
Begin VB.Form frmCompanyInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   1755
   ClientTop       =   615
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   12315
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3435
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   5280
      Width           =   5175
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   8040
      ScaleHeight     =   4395
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   5280
      Width           =   3855
   End
   Begin VB.CommandButton cmdSortRev 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort companies by revenue in descending order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortBeta 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort companies by Beta in ascending order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton CmdSortShrPrice 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort companies by share prices in descending order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdIndividualInfo 
      BackColor       =   &H0000FF00&
      Caption         =   "See individual company info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FF00&
      Caption         =   "Read."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.PictureBox picCompanies 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2760
      ScaleHeight     =   2955
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company Logos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "List of Companies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Query Results"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Various Companies and Related Financial Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   11775
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note: Revenue, EBITDA, and Market Cpaitalization figures are in millions of dollars."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   9
      Top             =   9000
      Width           =   5175
   End
End
Attribute VB_Name = "frmCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Project1(Financila_Instruments.vbp)
'Form: frmCompanyInfo(frmCompanyInfo.frm)
'Author: Sean Mase and David Horn
'Date Written: March 26, 2008
'Objective:  The purpose of this from is to give the user a list of companies and at the users request dipsplay.
            'company individual company info or sort all the companies under a certain data filed
            

Option Explicit

    'delcares variables glogal because they will be used more than one r more subroutine
    Dim Company(1 To 10) As String, Revenue(1 To 10) As Long, EBITDA(1 To 10) As Long
    Dim Ticker(1 To 10) As String, CompanyLogo(1 To 10) As String
    Dim MktCap(1 To 10) As Long, ShrPrice(1 To 10) As Single, PE(1 To 10) As Single
    Dim Beta(1 To 10) As Single, CTR As Integer, CTR2 As Integer
    Dim Pass As Integer, Pos As Integer, TempComp As String, TempTick As String
    Dim TempEBITDA As Long, TempRev As Long, TempMkt As Long, TempShr As Single
    Dim TempPE As Single, TempBeta As Single, TempLogo As String, C As Integer


Private Sub cmdIndividualInfo_Click()
    'This button allows the user display individual comany data
    
    'declares variables
    Dim NameInput As String, B As Integer, Found As Boolean

    'Asks user for company they would like to display
    NameInput = InputBox("Enter a company name or ticker symbol from the list above.")
    
    'delcares variable
    Found = False


    Do While (Not Found) And (B < CTR) 'do while seach that serarches until the item is found
        B = B + 1
        If NameInput = Company(B) Or NameInput = Ticker(B) Then Found = True 'makes boolean variable true if item is found
    Loop

    'displays one of two messages. One message if the item is found, the othe if the itrem is not found
    If Not Found Then
        MsgBox ("Sorry, the  company name or ticker symbol you entered is invalid." _
        & "Please enter another company name or ticker symbol.")
    Else
        picresults.Cls
        picresults.Print "Company: ", , Company(B)
        picresults.Print "Ticker Sybmol: ", Ticker(B)
        picresults.Print "**************************************************************"
        picresults.Print "Revenues:", , FormatCurrency(Revenue(B), 0)
        picresults.Print "EBITDA:", , FormatCurrency(EBITDA(B), 0)
        picresults.Print "Market Capitalization:", FormatCurrency(MktCap(B), 0)
        picresults.Print "Share Price:", , FormatCurrency(ShrPrice(B), 2)
        picresults.Print "P/E Raito:", , FormatNumber(PE(B), 2)
        picresults.Print "Beta:", , FormatNumber(Beta(B), 2)
        
        picLogo.Picture = LoadPicture(App.Path & "\" & CompanyLogo(B))  'displays the logo
    End If

End Sub

Private Sub cmdMainMenu_Click()
    'diplays main menu
    frmMainMenu.Show
    frmCompanyInfo.Hide
End Sub

Private Sub cmdRead_Click()
    'this button reads two files. One file gets the compnay information and the other get the company logos
    
    'delcares variable
    Dim A As Integer

    CTR = 0
    
    'opens file to be read into various arrays
    Open App.Path & "\Companies.txt" For Input As #1
    
    'loop to build arrays for the various company information
    Do Until EOF(1)     'loop until the loop reaches the end of the file
        CTR = CTR + 1   'increments counter
        Input #1, Company(CTR), Ticker(CTR), Revenue(CTR), EBITDA(CTR), MktCap(CTR), _
        ShrPrice(CTR), PE(CTR), Beta(CTR)   'builds the different arrays
    Loop
    
    'closes file #1
    Close #1
    
    'prints header for the list of companies in the file
    picCompanies.Cls
    picCompanies.Print "Companies", , "Ticker Symbol"
    picCompanies.Print "**************************************************************"
    
    'this for next statement properly spaces the company names and their ticker symbols
    For A = 1 To CTR
        If Len(Company(A)) >= 12 Then   'if the company name is greater than or equal to 12 characters do the spacing below
            picCompanies.Print Company(A), Ticker(A)
        Else
            picCompanies.Print Company(A), , Ticker(A)
        End If
    Next A

    CTR2 = 0
    
    'opens another file to be read into an array
    Open App.Path & "\CompanyLogos.txt" For Input As #2
    
    'loop to build arrays for the various company information
    Do Until EOF(2)         'loop until the loop reaches the end of the file
        CTR2 = CTR2 + 1     'increments counter
        Input #2, CompanyLogo(CTR2) 'builds the array
    Loop
    
    'closes file #2
    Close #2

End Sub

Private Sub cmdSortBeta_Click()
    'this button sorts all the companies into ascending order based on their Betas
    
    For Pass = 1 To CTR - 1     'Number of passes through the array
        For Pos = 1 To CTR - Pass       'number of comparisons through the array
             If Beta(Pos) > Beta(Pos + 1) Then
                TempComp = Company(Pos)     'switch the data if necessary
                TempTick = Ticker(Pos)
                TempRev = Revenue(Pos)
                TempEBITDA = EBITDA(Pos)
                TempMkt = MktCap(Pos)
                TempShr = ShrPrice(Pos)
                TempPE = PE(Pos)
                TempBeta = Beta(Pos)
                TempLogo = CompanyLogo(Pos)
                Company(Pos) = Company(Pos + 1)
                Ticker(Pos) = Ticker(Pos + 1)
                Revenue(Pos) = Revenue(Pos + 1)
                EBITDA(Pos) = EBITDA(Pos + 1)
                MktCap(Pos) = MktCap(Pos + 1)
                ShrPrice(Pos) = ShrPrice(Pos + 1)
                PE(Pos) = PE(Pos + 1)
                Beta(Pos) = Beta(Pos + 1)
                CompanyLogo(Pos) = CompanyLogo(Pos + 1)
                Company(Pos + 1) = TempComp
                Ticker(Pos + 1) = TempTick
                Revenue(Pos + 1) = TempRev
                EBITDA(Pos + 1) = TempEBITDA
                MktCap(Pos + 1) = TempMkt
                ShrPrice(Pos + 1) = TempShr
                PE(Pos + 1) = TempPE
                Beta(Pos + 1) = TempBeta
                CompanyLogo(Pos + 1) = TempLogo
            End If
        Next Pos
    Next Pass

    'prints the logo of the company that is at the top of the list
    picLogo.Cls
    
    'Prints header
    picresults.Cls
    picresults.Print "Company", , "Beta"
    picresults.Print "***************************************************************"
    
    'this for next statement properly spaces the company names and their ticker symbols
    For C = 1 To CTR
        If Len(Company(C)) >= 12 Then       'if the company name is greater than or equal to 12 characters do the spacing below
            picresults.Print Company(C), FormatNumber(Beta(C), 2)
            
        Else
            picresults.Print Company(C), , FormatNumber(Beta(C), 2)
        End If
    Next C
    
    'displays the logo of the company that is at the top of the list
    picLogo.Picture = LoadPicture(App.Path & "\" & CompanyLogo(C - 10))
End Sub

Private Sub cmdSortRev_Click()
    'This file sorts the companies into descending order based on their respective Revenue
    
    For Pass = 1 To CTR - 1     'Number of passes through the array
        For Pos = 1 To CTR - Pass       'number of comparisons through the array
             If Revenue(Pos) < Revenue(Pos + 1) Then
                TempComp = Company(Pos)     'switch the data if necessary
                TempTick = Ticker(Pos)
                TempRev = Revenue(Pos)
                TempEBITDA = EBITDA(Pos)
                TempMkt = MktCap(Pos)
                TempShr = ShrPrice(Pos)
                TempPE = PE(Pos)
                TempBeta = Beta(Pos)
                TempLogo = CompanyLogo(Pos)
                Company(Pos) = Company(Pos + 1)
                Ticker(Pos) = Ticker(Pos + 1)
                Revenue(Pos) = Revenue(Pos + 1)
                EBITDA(Pos) = EBITDA(Pos + 1)
                MktCap(Pos) = MktCap(Pos + 1)
                ShrPrice(Pos) = ShrPrice(Pos + 1)
                PE(Pos) = PE(Pos + 1)
                Beta(Pos) = Beta(Pos + 1)
                CompanyLogo(Pos) = CompanyLogo(Pos + 1)
                Company(Pos + 1) = TempComp
                Ticker(Pos + 1) = TempTick
                Revenue(Pos + 1) = TempRev
                EBITDA(Pos + 1) = TempEBITDA
                MktCap(Pos + 1) = TempMkt
                ShrPrice(Pos + 1) = TempShr
                PE(Pos + 1) = TempPE
                Beta(Pos + 1) = TempBeta
                CompanyLogo(Pos + 1) = TempLogo
            End If
        Next Pos
    Next Pass

    'prints the logo of the company that is at the top of the list
    picLogo.Cls
    
    'Prints header
    picresults.Cls
    picresults.Print "Company", , "Revenue"
    picresults.Print "***************************************************************"
    
    'this for next statement properly spaces the company names and their ticker symbols
    For C = 1 To CTR
        If Len(Company(C)) >= 12 Then       'if the company name is greater than or equal to 12 characters do the spacing below
            picresults.Print Company(C), FormatCurrency(Revenue(C), 0)
        Else
            picresults.Print Company(C), , FormatCurrency(Revenue(C), 0)
        End If
    Next C

    'displays the logo of the company that is at the top of the list
    picLogo.Picture = LoadPicture(App.Path & "\" & CompanyLogo(C - 10))
End Sub

Private Sub CmdSortShrPrice_Click()
    'This filr sorts the companies into descending order based on their repsective share prices
   
    
    For Pass = 1 To CTR - 1     'Number of passes through the array
        For Pos = 1 To CTR - Pass       'number of comparisons through the array
             If ShrPrice(Pos) < ShrPrice(Pos + 1) Then
                TempComp = Company(Pos)     'switch the data if necessary
                TempTick = Ticker(Pos)
                TempRev = Revenue(Pos)
                TempEBITDA = EBITDA(Pos)
                TempMkt = MktCap(Pos)
                TempShr = ShrPrice(Pos)
                TempPE = PE(Pos)
                TempBeta = Beta(Pos)
                TempLogo = CompanyLogo(Pos)
                Company(Pos) = Company(Pos + 1)
                Ticker(Pos) = Ticker(Pos + 1)
                Revenue(Pos) = Revenue(Pos + 1)
                EBITDA(Pos) = EBITDA(Pos + 1)
                MktCap(Pos) = MktCap(Pos + 1)
                ShrPrice(Pos) = ShrPrice(Pos + 1)
                PE(Pos) = PE(Pos + 1)
                Beta(Pos) = Beta(Pos + 1)
                CompanyLogo(Pos) = CompanyLogo(Pos + 1)
                Company(Pos + 1) = TempComp
                Ticker(Pos + 1) = TempTick
                Revenue(Pos + 1) = TempRev
                EBITDA(Pos + 1) = TempEBITDA
                MktCap(Pos + 1) = TempMkt
                ShrPrice(Pos + 1) = TempShr
                PE(Pos + 1) = TempPE
                Beta(Pos + 1) = TempBeta
                CompanyLogo(Pos + 1) = TempLogo
            End If
        Next Pos
    Next Pass

    'prints the logo of the company that is at the top of the list
    picLogo.Cls
    
    'Prints Header
    picresults.Cls
    picresults.Print "Company", , "Share Price"
    picresults.Print "***************************************************************"
    
    'this for next statement properly spaces the company names and their ticker symbols
    For C = 1 To CTR
        If Len(Company(C)) >= 12 Then       'if the company name is greater than or equal to 12 characters do the spacing below
            picresults.Print Company(C), FormatCurrency(ShrPrice(C), 2)
        Else
            picresults.Print Company(C), , FormatCurrency(ShrPrice(C), 2)
        End If
    Next C

    'displays the logo of the company that is at the top of the list
    picLogo.Picture = LoadPicture(App.Path & "\" & CompanyLogo(C - 10))
End Sub

Private Sub Form_Load()
    'displays message before the form opens
    MsgBox ("CLICK THE 'READ' BUTTON FIRST!")
    
End Sub

