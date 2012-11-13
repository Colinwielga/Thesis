Attribute VB_Name = "Module1"
'all of these variable are used throughout many forms on our project

Public WestTotal As Integer     'total score for west
Public SouthTotal As Integer     'total score for south
Public EastTotal As Integer        'total score for east
Public MidwestTotal As Integer      'total score for midwest
Public FinalsTotal As Integer          'final four total
Public WestWinner As String                'winner of west region
Public EastWinner As String                 'winner of east region
Public SouthWinner As String                'winner of south region
Public MidwestWinner As String              'winner of midwest region
Public User As String                       'user name
Public SouthR1Sum As Integer, SouthR2Sum As Integer, SouthR3Sum As Integer      'next four lines are the sums for each region and each round
Public EastR1Sum As Integer, EastR2Sum As Integer, EastR3Sum As Integer
Public WestR1Sum As Integer, WestR2Sum As Integer, WestR3Sum As Integer
Public MidwestR1Sum As Integer, MidwestR2Sum As Integer, MidwestR3Sum As Integer
Public Final4Sum As Integer, ChampionshipSum As Integer, ChampsSum As Integer
Public OverallScore As Integer          'overall score for entire bracket

