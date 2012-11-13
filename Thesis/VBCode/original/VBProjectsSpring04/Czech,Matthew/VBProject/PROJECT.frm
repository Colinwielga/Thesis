VERSION 5.00
Begin VB.Form frmstatscalsulations 
   BackColor       =   &H000000FF&
   Caption         =   "Leaders"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdallstats 
      Caption         =   "Click Here For Statistics of all SJU HOCKEY PLAYERS Of 2003-2004"
      Height          =   5295
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdpoints 
      Caption         =   "Click to get The Top 10 Point Getters"
      Height          =   1455
      Left            =   5400
      TabIndex        =   6
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CommandButton cmdassists 
      Caption         =   "Click to get The Top 10 Playmakers"
      Height          =   1455
      Left            =   2760
      TabIndex        =   5
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Click here once done"
      Height          =   2655
      Left            =   13440
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdpenmin 
      Caption         =   "Click here to get the top 10 penalty minute leaders"
      Height          =   1455
      Left            =   10680
      TabIndex        =   3
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CommandButton cmdplusmin 
      Caption         =   "Click here to get The Best 10 Plus/Minus"
      Height          =   1455
      Left            =   8040
      TabIndex        =   2
      Top             =   9360
      Width           =   2535
   End
   Begin VB.CommandButton cmdgoalleader 
      BackColor       =   &H00FF8080&
      Caption         =   "Click to get The Top 10 Goal Scorers"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   9360
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      Height          =   5175
      Left            =   2400
      ScaleHeight     =   5115
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   120
      Width           =   12855
   End
   Begin VB.Label lblregchamps 
      Caption         =   "2003-2004 Regular Season Conference Champions"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   6000
      Width           =   3855
   End
   Begin VB.Image Imgchamps 
      Height          =   2535
      Left            =   5160
      Picture         =   "PROJECT.frx":0000
      Top             =   6480
      Width           =   5400
   End
   Begin VB.Label creator 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Czech"
      Height          =   255
      Left            =   10560
      TabIndex        =   7
      Top             =   7920
      Width           =   1695
   End
End
Attribute VB_Name = "frmstatscalsulations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : SJU HOCKEY (Matthew Czech's VB Project.vbp)
'Form Name : frmstatscalsulations (PROJECT.frm)
'Author: Matthew Czech
'Date Written: March 12, 2003
'Purpose of Form: To Calculate and sort given data to show the best in certain categories
                    ' Also to print the data used.

Private Sub cmdallstats_Click()
    picresults.Cls
    picresults.Print "Jersey #'s", "Names", , "Played", "Goals", "Assists", "Shots", "plus/min", "Penalty Min.", "Power-Play", "Short-Handed"; "Game Winners"
    picresults.Print "_________________________________________________________________________________________________________________________________________________________"
        For j = 1 To 45 'fills arrays
            picresults.Print numbers(j), names(j), gp(j), goals(j), assists(j), shots(j), plusmin(j), penmin(j), pp(j), sh(j), gw(j) 'prints arrays
        Next j
End Sub

Private Sub cmdassists_Click()
    picresults.Cls
    CTR = 45
    For Pass = 1 To (CTR - 1) 'to sort arrays my number fo assists from larget to smalles
        For Comp = 1 To (CTR - 1)
            If assists(Comp) < assists(Comp + 1) Then
            
                tempgoals = goals(Comp)         'Following keeps arrays in proper order"d
                goals(Comp) = goals(Comp + 1)
                goals(Comp + 1) = tempgoals
            
                tempnumbers = numbers(Comp)
                numbers(Comp) = numbers(Comp + 1)
                numbers(Comp + 1) = tempnumbers
                
                tempnames = names(Comp)
                names(Comp) = names(Comp + 1)
                names(Comp + 1) = tempnames
    
                tempassists = assists(Comp)
                assists(Comp) = assists(Comp + 1)
                assists(Comp + 1) = tempassists
    
                tempgp = gp(Comp)
                gp(Comp) = gp(Comp + 1)
                    gp(Comp + 1) = gpnames
            
                    temppoints = points(Comp)
                    points(Comp) = points(Comp + 1)
                    points(Comp + 1) = temppoints
                    
                tempshots = shots(Comp)
                shots(Comp) = shots(Comp + 1)
                shots(Comp + 1) = tempshots
        
                    tempshtpct = shtpct(Comp)
                shtpct(Comp) = shtpct(Comp + 1)
                shtpct(Comp + 1) = tempshtpct
        
                tempplusmin = plusmin(Comp)
                plusmin(Comp) = plusmin(Comp + 1)
                plusmin(Comp + 1) = tempplusmin
                
                temppenmin = penmin(Comp)
                penmin(Comp) = penmin(Comp + 1)
                penmin(Comp + 1) = temppenmin
        
                temppp = pp(Comp)
                pp(Comp) = pp(Comp + 1)
                pp(Comp + 1) = temppp
        
                tempsh = sh(Comp)
                sh(Comp) = sh(Comp + 1)
                sh(Comp + 1) = tempsh
                
                tempgw = gw(Comp)
                gw(Comp) = gw(Comp + 1)
                gw(Comp + 1) = tempgw
            End If
        Next Comp
    Next Pass
    picresults.Print "The Top 10 Assist makers for your St. John's Hockey ARE:"
    picresults.Print "*********************************************************************"
    picresults.Print "Jersey #", "Name", , "Assists"
    picresults.Print "--------------------------------------------------------------------------------------------------------------------"
    For j = 1 To 10
        picresults.Print numbers(j), names(j); , assists(j)
    Next j
    cmdgoalleader.Enabled = False
    cmdplusmin.Enabled = False
    cmdpenmin.Enabled = False
    cmdassists.Enabled = False
    cmdpoints.Enabled = True
End Sub


Private Sub cmdgoalleader_Click()
'prints to ten goal scroers and keeps them with the correct data
        CTR = 45
    picresults.Cls
    For Pass = 1 To (CTR - 1)
        For Comp = 1 To (CTR - 1)
            If goals(Comp) < goals(Comp + 1) Then
            
                tempgoals = goals(Comp)
                goals(Comp) = goals(Comp + 1)
                goals(Comp + 1) = tempgoals
            
                tempnumbers = numbers(Comp)
                numbers(Comp) = numbers(Comp + 1)
                numbers(Comp + 1) = tempnumbers
                
                tempnames = names(Comp)
                names(Comp) = names(Comp + 1)
                names(Comp + 1) = tempnames
        
                tempassists = assists(Comp)
                assists(Comp) = assists(Comp + 1)
                assists(Comp + 1) = tempassists
        
                tempgp = gp(Comp)
                gp(Comp) = gp(Comp + 1)
                gp(Comp + 1) = gpnames
        
                temppoints = points(Comp)
                points(Comp) = points(Comp + 1)
                points(Comp + 1) = temppoints
                
                tempshots = shots(Comp)
                shots(Comp) = shots(Comp + 1)
                shots(Comp + 1) = tempshots
        
                tempshtpct = shtpct(Comp)
                shtpct(Comp) = shtpct(Comp + 1)
                shtpct(Comp + 1) = tempshtpct
        
                tempplusmin = plusmin(Comp)
                plusmin(Comp) = plusmin(Comp + 1)
                plusmin(Comp + 1) = tempplusmin
                
                temppenmin = penmin(Comp)
                penmin(Comp) = penmin(Comp + 1)
                penmin(Comp + 1) = temppenmin
        
                temppp = pp(Comp)
                pp(Comp) = pp(Comp + 1)
                pp(Comp + 1) = temppp
        
                tempsh = sh(Comp)
                sh(Comp) = sh(Comp + 1)
                sh(Comp + 1) = tempsh
                
                tempgw = gw(Comp)
                gw(Comp) = gw(Comp + 1)
                gw(Comp + 1) = tempgw
            
            End If
        Next Comp
    Next Pass
    picresults.Print "The Top 10 Goal Scorers for your St. John's Hockey ARE:"
    picresults.Print "*********************************************************************"
    picresults.Print "Jersey #", "Name", , "Goals"
    picresults.Print "--------------------------------------------------------------------------------------------------------------------"
    For j = 1 To 10
        picresults.Print numbers(j), names(j); , goals(j)
    Next j
    cmdgoalleader.Enabled = False
    cmdplusmin.Enabled = False
    cmdpenmin.Enabled = False
    cmdpoints.Enabled = False
    cmdassists.Enabled = True
End Sub

Private Sub cmdpenmin_Click()
    picresults.Cls
    CTR = 45
    For Pass = 1 To (CTR - 1) 'To sort according to Penalty Minutes
        For Comp = 1 To (CTR - 1)
            If penmin(Comp) < penmin(Comp + 1) Then
            
                tempgoals = goals(Comp)
                goals(Comp) = goals(Comp + 1)
                goals(Comp + 1) = tempgoals
                
                tempnumbers = numbers(Comp)
                numbers(Comp) = numbers(Comp + 1)
                numbers(Comp + 1) = tempnumbers
                
                tempnames = names(Comp)
                names(Comp) = names(Comp + 1)
                names(Comp + 1) = tempnames
        
    
                
                tempassists = assists(Comp)
                assists(Comp) = assists(Comp + 1)
                assists(Comp + 1) = tempassists
        
                tempgp = gp(Comp)
                gp(Comp) = gp(Comp + 1)
                gp(Comp + 1) = gpnames
        
                temppoints = points(Comp)
                points(Comp) = points(Comp + 1)
                points(Comp + 1) = temppoints
                
                tempshots = shots(Comp)
                shots(Comp) = shots(Comp + 1)
                shots(Comp + 1) = tempshots
        
                tempshtpct = shtpct(Comp)
                shtpct(Comp) = shtpct(Comp + 1)
                shtpct(Comp + 1) = tempshtpct
        
                tempplusmin = plusmin(Comp)
                plusmin(Comp) = plusmin(Comp + 1)
                plusmin(Comp + 1) = tempplusmin
                
                temppenmin = penmin(Comp)
                penmin(Comp) = penmin(Comp + 1)
                penmin(Comp + 1) = temppenmin
        
                temppp = pp(Comp)
                pp(Comp) = pp(Comp + 1)
                pp(Comp + 1) = temppp
        
                tempsh = sh(Comp)
                sh(Comp) = sh(Comp + 1)
                sh(Comp + 1) = tempsh
                
                tempgw = gw(Comp)
                gw(Comp) = gw(Comp + 1)
                gw(Comp + 1) = tempgw
            End If
        Next Comp
    Next Pass

    picresults.Print "The Top 10 Penalty guys for your St. John's Hockey ARE:"
    picresults.Print "**********************************************************************"
    picresults.Print "Jersey #", "Name", , "Penalty Minutes"
    picresults.Print "--------------------------------------------------------------------------------------------------------------------"
    For j = 1 To 10
       picresults.Print numbers(j), names(j); , penmin(j)
    Next j
    cmdgoalleader.Enabled = True
    cmdplusmin.Enabled = True
    cmdpenmin.Enabled = True
    cmdassists.Enabled = True
    cmdpoints.Enabled = True
    cmdquit.Enabled = True
End Sub
Private Sub cmdplusmin_Click()
    picresults.Cls
    CTR = 45
    For Pass = 1 To (CTR - 1) 'To sort Plus/Minus
        For Comp = 1 To (CTR - 1)
            If plusmin(Comp) < plusmin(Comp + 1) Then
                
            tempgoals = goals(Comp)
            goals(Comp) = goals(Comp + 1)
            goals(Comp + 1) = tempgoals
            
            tempnumbers = numbers(Comp)
            numbers(Comp) = numbers(Comp + 1)
            numbers(Comp + 1) = tempnumbers
            
            tempnames = names(Comp)
            names(Comp) = names(Comp + 1)
            names(Comp + 1) = tempnames
    

            
            tempassists = assists(Comp)
            assists(Comp) = assists(Comp + 1)
            assists(Comp + 1) = tempassists
    
            tempgp = gp(Comp)
            gp(Comp) = gp(Comp + 1)
            gp(Comp + 1) = gpnames
    
            temppoints = points(Comp)
            points(Comp) = points(Comp + 1)
            points(Comp + 1) = temppoints
            
            tempshots = shots(Comp)
            shots(Comp) = shots(Comp + 1)
            shots(Comp + 1) = tempshots
    
            tempshtpct = shtpct(Comp)
            shtpct(Comp) = shtpct(Comp + 1)
            shtpct(Comp + 1) = tempshtpct
    
            tempplusmin = plusmin(Comp)
            plusmin(Comp) = plusmin(Comp + 1)
            plusmin(Comp + 1) = tempplusmin
            
            temppenmin = penmin(Comp)
            penmin(Comp) = penmin(Comp + 1)
            penmin(Comp + 1) = temppenmin
    
            temppp = pp(Comp)
            pp(Comp) = pp(Comp + 1)
            pp(Comp + 1) = temppp
    
            tempsh = sh(Comp)
            sh(Comp) = sh(Comp + 1)
            sh(Comp + 1) = tempsh
            
            tempgw = gw(Comp)
            gw(Comp) = gw(Comp + 1)
            gw(Comp + 1) = tempgw
        End If
    Next Comp
Next Pass
    picresults.Print "The Top 10 Plus/Minus Rating for your St. John's Hockey ARE:"
    picresults.Print "*********************************************************************"
    picresults.Print "Jersey #", "Name", , "Plus/Minus"
    picresults.Print "--------------------------------------------------------------------------------------------------------------------"
    For j = 1 To 10
        picresults.Print numbers(j), names(j); , plusmin(j)
    Next j
    cmdgoalleader.Enabled = False
    cmdplusmin.Enabled = False
    cmdpoints.Enabled = False
    cmdpenmin.Enabled = True

End Sub



Private Sub cmdpoints_Click()
picresults.Cls
CTR = 45                        'Calculates points
    For Pass = 1 To (CTR - 1)   'Sorts by points
        For Comp = 1 To (CTR - 1)
          points(Comp) = goals(Comp) + assists(Comp)
                If points(Comp) < points(Comp + 1) Then
                    temppoints = points(Comp)
                    points(Comp) = points(Comp + 1)
                    points(Comp + 1) = temppoints
                
                    tempgoals = goals(Comp)
                    goals(Comp) = goals(Comp + 1)
                    goals(Comp + 1) = tempgoals
            
                    tempnumbers = numbers(Comp)
                    numbers(Comp) = numbers(Comp + 1)
                    numbers(Comp + 1) = tempnumbers
            
                    tempnames = names(Comp)
                    names(Comp) = names(Comp + 1)
                    names(Comp + 1) = tempnames
    

            
                    tempassists = assists(Comp)
                    assists(Comp) = assists(Comp + 1)
                    assists(Comp + 1) = tempassists
    
                    tempgp = gp(Comp)
                    gp(Comp) = gp(Comp + 1)
                    gp(Comp + 1) = gpnames
    
                    tempshots = shots(Comp)
                    shots(Comp) = shots(Comp + 1)
                    shots(Comp + 1) = tempshots
    
                    tempshtpct = shtpct(Comp)
                    shtpct(Comp) = shtpct(Comp + 1)
                    shtpct(Comp + 1) = tempshtpct
    
                    tempplusmin = plusmin(Comp)
                    plusmin(Comp) = plusmin(Comp + 1)
                    plusmin(Comp + 1) = tempplusmin
            
                    temppenmin = penmin(Comp)
                    penmin(Comp) = penmin(Comp + 1)
                    penmin(Comp + 1) = temppenmin
    
                    temppp = pp(Comp)
                    pp(Comp) = pp(Comp + 1)
                    pp(Comp + 1) = temppp
    
                    tempsh = sh(Comp)
                    sh(Comp) = sh(Comp + 1)
                    sh(Comp + 1) = tempsh
            
                    tempgw = gw(Comp)
                    gw(Comp) = gw(Comp + 1)
                    gw(Comp + 1) = tempgw
End If
    Next Comp
    Next Pass
    picresults.Print "The Top 10 Point getters for your St. John's Hockey ARE:"
    picresults.Print "*********************************************************************"
    picresults.Print "Jersey #", "Name", , "Points"
    picresults.Print "--------------------------------------------------------------------------------------------------------------------"
    For j = 1 To 10
        picresults.Print numbers(j), names(j); , points(j)
    Next j
    cmdgoalleader.Enabled = False
    cmdplusmin.Enabled = True
    cmdpoints.Enabled = False
    cmdpenmin.Enabled = False
    cmdquit.Enabled = True

End Sub

Private Sub cmdquit_Click()
    MsgBox "You can't leave just yet, You must check out our online store", , "20% OFF on your first online purchase!!!"
    frmstatscalsulations.Hide
    frmgiftshop.Show 'Next form
End Sub


Private Sub Form_Load()
cmdgoalleader.Enabled = True
cmdplusmin.Enabled = False
cmdassists.Enabled = False
cmdpenmin.Enabled = False
cmdpoints.Enabled = False
cmdquit.Enabled = True

End Sub

