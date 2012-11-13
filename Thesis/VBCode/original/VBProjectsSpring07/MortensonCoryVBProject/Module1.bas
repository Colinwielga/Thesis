Attribute VB_Name = "Module1"
'This code module declares variables for all forms
    Public AGI As Double, N As String
    Public Exemptions As Integer, Deduction As Long
    Public TaxableIncome As Double
    Public TaxLiability As Single, TaxesWithheld As Double
    Public Refund As Double, Payment As Double
    
    Public Bracket(1 To 100) As Double, Risk(1 To 100) As String, Potential(1 To 100) As Single
    
    Public CTR1 As Single, Pos As Single
    
    Public Medical As Single, Taxes As Single, Interest As Single, Gifts As Single, Loss As Single, Job As Single, Misc As Single, Itemized As Single
    Public Deduct As Double, Standard As Single
    Public Children As Integer, Credit As Single
    
    Public ClientArray(1 To 1000) As String, AGIArray(1 To 1000) As Double, IncomeArray(1 To 1000) As Double, LiabilityArray(1 To 1000) As Double, WithheldArray(1 To 1000) As Double
    
    
