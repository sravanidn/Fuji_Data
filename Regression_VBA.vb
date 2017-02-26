
Private Sub CommandButton1_Click()
' This sub computes the prediction intervals for the 7 day data entered in the Userform
Dim tstat As Double
Dim dt As Date
Dim g4 As Integer
Dim x(8, 4) As String
Dim h4 As Integer
Dim i4 As Integer
Dim j4 As Integer
h4 = 0
i4 = 0
j4 = 0
g4 = 125
Dim region As String
If ListBox1.ListCount < 8 Then
 MsgBox " Kindly add 7 day data for computing forecast"
 Exit Sub
End If
UserForm3.Show
region = Label3.Caption
For I = 1 To ListBox1.ListCount - 1
    For J = 0 To 3
    x(I, J) = ListBox1.List(I, J)
    Next
Next
'BLOOMFIELD
If region = "Bloomfield" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Bloomfield-Regression").Range("J" & J)
    predicted = Sheets("Bloomfield-Regression").Range("K" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
'degreefreedom = 125 - 4 - 1
degreefreedom = 125 - 3 - 1


tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
'MsgBox 'T-Stat:' tstat
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 1) = "Warm" Then
        h4 = 1
    End If
   ' If x(I, 1) = "Cool" Then
    '    i4 = 1
    'End If
    'If x(I, 2) = "Dry" Then
    '    j4 = 1
    'End If
    'If x(I, 1) = "Very Dry" Then
    '    k4 = 1
    'End If
    If x(I, 3) = 1 Then
        l4 = 1
    End If

prediction = Sheets("Bloomfield-Regression").Range("V19") + Sheets("Bloomfield-Regression").Range("V20") * g4 + Sheets("Bloomfield-Regression").Range("V21") * l4

'calculating interval

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Bloomfield"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
'MsgBox prediction
Next


'Farmington - $V$19+$V$20*G4+$V$21*H4+$V$22*I4+$V$23*J4+$V$24*K4+$V$25*L4
ElseIf region = "Farmington" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Farmington-Regression").Range("M" & J)
    predicted = Sheets("Farmington-Regression").Range("N" & J)
    rse = (actual - predicted) ^ 2 + rse
    
Next
rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 6 - 1
tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
    
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 1) = "Mild" Then
        h4 = 1
    End If
    If x(I, 1) = "Warm" Then
        i4 = 1
    End If
        If x(I, 2) = "Dry" Then
        j4 = 1
    End If
        If x(I, 2) = "Very dry" Then
        k4 = 1
    End If
    If x(I, 3) = 1 Then
        l4 = 1
    End If

prediction = (Sheets("Farmington-Regression").Range("V19") + _
             Sheets("Farmington-Regression").Range("V20") * g4 + _
             Sheets("Farmington-Regression").Range("V21") * h4 + _
             Sheets("Farmington-Regression").Range("V22") * i4 + _
             Sheets("Farmington-Regression").Range("V23") * j4 + _
             Sheets("Farmington-Regression").Range("V24") * k4 + _
             Sheets("Farmington-Regression").Range("V25") * l4)
'MsgBox 'T-Stat:' tstat

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Farmington"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next

'Glastonbury =$V$19+$V$20*G4+$V$21*H4+$V$22*I4+$V$23*J4+$V$24*K4+$V$25*L4
ElseIf region = "Glastonbury" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Glastonbury-Regression").Range("L" & J)
    predicted = Sheets("Glastonbury-Regression").Range("M" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 5 - 1



tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
  '  If x(I, 1) = "Mild" Then
  '      h4 = 1
  '  End If
    If x(I, 1) = "Warm" Then
        h4 = 1
    End If
    If x(I, 1) = "Cool" Then
        i4 = 1
    End If
    If x(I, 2) = "Dry" Then
        j4 = 1
    End If
    If x(I, 3) = 1 Then
        k4 = 1
    End If

prediction = (Sheets("Glastonbury-Regression").Range("X19") + _
Sheets("Glastonbury-Regression").Range("X20") * g4 + _
Sheets("Glastonbury-Regression").Range("X21") * h4 + _
Sheets("Glastonbury-Regression").Range("X22") * i4 + _
Sheets("Glastonbury-Regression").Range("X23") * j4 + _
Sheets("Glastonbury-Regression").Range("X24") * k4)


'MsgBox 'T-Stat:' tstat

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Glastonbury"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next
'Hartford - =$V$19+$V$20*G4+$V$21*H4+$V$22*I4+$V$23*J4+$V$24*K4+$V$25*L4
ElseIf region = "Hartford" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Hartford - Regression").Range("M" & J)
    predicted = Sheets("Hartford - Regression").Range("N" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 6 - 1

tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)

For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 1) = "Mild" Then
        h4 = 1
    End If
    If x(I, 1) = "Cool" Then
        i4 = 1
    End If
    If x(I, 2) = "Dry" Then
        j4 = 1
    End If
    If x(I, 2) = "Very Dry" Then
        k4 = 1
    End If
    If x(I, 3) = 1 Then
        l4 = 1
    End If

prediction = Sheets("Hartford - Regression").Range("Y19") + Sheets("Hartford - Regression").Range("Y20") * g4 + Sheets("Hartford - Regression").Range("Y21") * h4 + Sheets("Hartford - Regression").Range("Y22") * i4 + Sheets("Hartford - Regression").Range("Y23") * j4 + Sheets("Hartford - Regression").Range("Y24") * k4 + Sheets("Hartford - Regression").Range("Y25") * l4


'MsgBox 'T-Stat:' tstat

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Hartford"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next
' Manchester-Regression  =$R$19+$R$20*G4+$R$21*H4

ElseIf region = "Manchester" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Manchester-Regression").Range("I" & J)
    predicted = Sheets("Manchester-Regression").Range("J" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 2 - 1

tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 3) = 1 Then
        h4 = 1
    End If

prediction = Sheets("Manchester-Regression").Range("U19") + Sheets("Manchester-Regression").Range("U20") * g4 + Sheets("Manchester-Regression").Range("U21") * h4

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Manchester"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next
'Southington-Regression $V$19+$V$20*G4+$V$21*H4+$V$22*I4+$V$23*J4+$V$24*K4+$V$25*L4
ElseIf region = "Southington" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Southington-Regression").Range("M" & J)
    predicted = Sheets("Southington-Regression").Range("N" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 6 - 1

tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 1) = "Mild" Then
        h4 = 1
    End If
    If x(I, 1) = "Warm" Then
        i4 = 1
    End If
    If x(I, 2) = "Dry" Then
        j4 = 1
    End If
    If x(I, 2) = "Very dry" Then
        k4 = 1
    End If
    If x(I, 3) = 1 Then
        l4 = 1
    End If

prediction = Sheets("Southington-Regression").Range("Y19") + Sheets("Southington-Regression").Range("Y20") * g4 + Sheets("Southington-Regression").Range("Y21") * h4 + Sheets("Southington-Regression").Range("Y22") * i4 + Sheets("Southington-Regression").Range("Y23") * j4 + Sheets("Southington-Regression").Range("Y24") * k4 + Sheets("Southington-Regression").Range("Y25") * l4

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Southington"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next
'West Hartford-Regression $W$19+$W$20*G4+$W$21*H4+$W$22*I4+$W$23*J4+$W$24*K4+$W$25*L4+$W$26*M4
ElseIf region = "West Hartford" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("West Hartford-Regression").Range("N" & J)
    predicted = Sheets("West Hartford-Regression").Range("O" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 7 - 1

tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    If x(I, 1) = "Mild" Then
        h4 = 1
    End If
    If x(I, 1) = "Cool" Then
        i4 = 1
    End If
    If x(I, 2) = "Dry" Then
        j4 = 1
    End If
    If x(I, 2) = "Humid" Then
        k4 = 1
    End If
    If x(I, 2) = "Very dry" Then
        l4 = 1
    End If
    If x(I, 3) = 1 Then
        m4 = 1
    End If

prediction = Sheets("West Hartford-Regression").Range("Z19") + Sheets("West Hartford-Regression").Range("Z20") * g4 + Sheets("West Hartford-Regression").Range("Z21") * h4 + Sheets("West Hartford-Regression").Range("Z22") * i4 + Sheets("West Hartford-Regression").Range("Z23") * j4 + Sheets("West Hartford-Regression").Range("Z24") * k4 + Sheets("West Hartford-Regression").Range("Z26") * l4 + Sheets("West Hartford-Regression").Range("Z26") * m4

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "West Hartford"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next
'Windsor-Regression =$X$19+$X$20*G12+$X$21*H12+$X$22*I12+$X$23*J12+$X$24*K12
ElseIf region = "Windsor" Then
rse = 0
For J = 4 To 128
    
    actual = Sheets("Windsor-Regression").Range("L" & J)
    predicted = Sheets("Windsor-Regression").Range("M" & J)
    rse = (actual - predicted) ^ 2 + rse
    
    
Next

rse = rse / 125
rse = Sqr(rse)
'MsgBox "RSE " & rse
siglevel = 0.05
degreefreedom = 125 - 5 - 1

tstat = Application.WorksheetFunction.T_Inv_2T(siglevel, degreefreedom)
For I = 1 To ListBox1.ListCount - 1
    dt = ListBox1.List(I, 0)
    g4 = g4 + 1
    'If x(I, 1) = "Mild" Then
    '    h4 = 1
    'End If
    'If x(I, 1) = "Cool" Then
    '    i4 = 1
    'End If
    If x(I, 2) = "Dry" Then
        h4 = 1
    End If
    If x(I, 2) = "Humid" Then
        i4 = 1
    End If
    If x(I, 2) = "Very Dry" Then
        j4 = 1
    End If
    If x(I, 3) = 1 Then
        k4 = 1
    End If
prediction = Sheets("Windsor-Regression").Range("X19") + Sheets("Windsor-Regression").Range("X20") * g4 + Sheets("Windsor-Regression").Range("X21") * h4 + Sheets("Windsor-Regression").Range("X22") * i4 + Sheets("Windsor-Regression").Range("X23") * j4 + Sheets("Windsor-Regression").Range("X24") * k4

lowerbound = prediction - rse * tstat
upperbound = prediction + rse * tstat


'adding data to new userform
UserForm4.ListBox1.AddItem
With UserForm4
      .ListBox1.List(I, 0) = "Windsor"
      .ListBox1.List(I, 1) = dt
      .ListBox1.List(I, 2) = prediction
      .ListBox1.List(I, 3) = lowerbound
      .ListBox1.List(I, 4) = upperbound
End With
Next


End If

UserForm4.Show
'MsgBox prediction

End Sub

Private Sub CommandButton2_Click()
Dim temp As String ' temperature
Dim hmd As String   ' humidity'
Dim fday As Integer ' festive day

temp = ComboBox1.Value
hmd = ComboBox2.Value
If CheckBox1.Value = True Then
    fday = 1
Else
    fday = 0
End If
If ListBox1.ListCount < 8 Then
    ListBox1.AddItem
    cnt = ListBox1.ListCount - 1
     ListBox1.List(cnt, 0) = Label5.Caption
    ListBox1.List(cnt, 1) = temp
    ListBox1.List(cnt, 2) = hmd
    ListBox1.List(cnt, 3) = fday
Else
    MsgBox " Already 7 day data is there, more data can not be entered"
    Exit Sub
End If
Label5.Caption = ListBox1.List(ListBox1.ListCount - 1, 0)
Label5.Caption = DateAdd("d", 1, Label5.Caption)
End Sub

Private Sub CommandButton3_Click()
If ListBox1.ListIndex > 0 Then
I = ListBox1.ListCount - 1
    While I >= 0
       If ListBox1.Selected(I) Then
       If I = ListBox1.ListCount - 1 Then
           ListBox1.RemoveItem (I)
           Label5.Caption = DateAdd("d", 1, ListBox1.List(ListBox1.ListCount - 1, 0))
           Exit Sub
           
           Else
           ListBox1.RemoveItem (I)
           End If
           
           ListBox1.Selected(I) = False
       End If
        I = I - 1
    Wend
End If
End Sub

Private Sub CommandButton4_Click()
' This sub is to Revise Models
Dim numrows As Integer
Dim Yrange As Range
Dim Xrange As Range
Dim outputrange As Range
Dim seasonalfactor As Double

Dim criteria As Range
Dim sumrange As Range
Dim output As String

numrows = 128
totalrows = Sheets("Orders Summary").Range("B3").End(xlDown).Row
numrows = numrows + 1

For I = numrows To totalrows

Set ordersum = ThisWorkbook.Sheets("Orders Summary")

newdate = ordersum.Cells(numrows, 2).Value
newday = ordersum.Cells(numrows, 3).Value
newtemp = ordersum.Cells(numrows, 4).Value
newhum = ordersum.Cells(numrows, 5).Value
newfest = ordersum.Cells(numrows, 6).Value
newbloomfield = ordersum.Cells(numrows, 7).Value
newfarmington = ordersum.Cells(numrows, 8).Value
newglaston = ordersum.Cells(numrows, 9).Value
newhartford = ordersum.Cells(numrows, 10).Value
newmanchester = ordersum.Cells(numrows, 11).Value
newsouthington = ordersum.Cells(numrows, 12).Value
newwesthartford = ordersum.Cells(numrows, 13).Value
newwindsor = ordersum.Cells(numrows, 14).Value

'Bloomfield
Set bloomfield = ThisWorkbook.Sheets("Bloomfield-Regression")
bloomfield.Cells(numrows, 2).Value = newdate
bloomfield.Cells(numrows, 3).Value = newday
bloomfield.Cells(numrows, 4).Value = newtemp
bloomfield.Cells(numrows, 5).Value = newhum
bloomfield.Cells(numrows, 6).Value = newfest
bloomfield.Cells(numrows, 7).Value = bloomfield.Cells(numrows - 1, 7).Value + 1
If newtemp = "Warm" Then
bloomfield.Cells(numrows, 8).Value = 1
Else
bloomfield.Cells(numrows, 8).Value = 0
End If
'If newtemp = "Cool" Then
'bloomfield.Cells(numrows, 9).Value = 1
'Else
'bloomfield.Cells(numrows, 9).Value = 0
'End If
If newfest = "Yes" Then
bloomfield.Cells(numrows, 9).Value = 1
Else
bloomfield.Cells(numrows, 9).Value = 0
End If
bloomfield.Cells(numrows, 10).Value = newbloomfield

'Farmington
Set Farmington = ThisWorkbook.Sheets("Farmington-Regression")
Farmington.Cells(numrows, 2).Value = newdate
Farmington.Cells(numrows, 3).Value = newday
Farmington.Cells(numrows, 4).Value = newtemp
Farmington.Cells(numrows, 5).Value = newhum
Farmington.Cells(numrows, 6).Value = newfest
Farmington.Cells(numrows, 7).Value = Farmington.Cells(numrows - 1, 7).Value + 1
If newtemp = "Mild" Then
Farmington.Cells(numrows, 8).Value = 1
Else
Farmington.Cells(numrows, 8).Value = 0
End If
If newtemp = "Warm" Then
Farmington.Cells(numrows, 9).Value = 1
Else
Farmington.Cells(numrows, 9).Value = 0
End If
If newhum = "Dry" Then
Farmington.Cells(numrows, 10).Value = 1
Else
Farmington.Cells(numrows, 10).Value = 0
End If
If newhum = "Very Dry" Then
Farmington.Cells(numrows, 11).Value = 1
Else
Farmington.Cells(numrows, 11).Value = 0
End If

If newfest = "Yes" Then
Farmington.Cells(numrows, 12).Value = 1
Else
Farmington.Cells(numrows, 12).Value = 0
End If
Farmington.Cells(numrows, 13).Value = newfarmington

'Glastonbury
Sheets("Glastonbury-Regression").Cells(numrows, 2).Value = newdate
Sheets("Glastonbury-Regression").Cells(numrows, 3).Value = newday
Sheets("Glastonbury-Regression").Cells(numrows, 4).Value = newtemp
Sheets("Glastonbury-Regression").Cells(numrows, 5).Value = newhum
Sheets("Glastonbury-Regression").Cells(numrows, 6).Value = newfest
Sheets("Glastonbury-Regression").Cells(numrows, 7).Value = Sheets("Glastonbury-Regression").Cells(numrows - 1, 7).Value + 1
'If newtemp = "Mild" Then
'Sheets("Glastonbury-Regression").Cells(numrows, 8).Value = 1
'Else
'Sheets("Glastonbury-Regression").Cells(numrows, 8).Value = 0
'End If
If newtemp = "Warm" Then
Sheets("Glastonbury-Regression").Cells(numrows, 8).Value = 1
Else
Sheets("Glastonbury-Regression").Cells(numrows, 8).Value = 0
End If
If newtemp = "Cool" Then
Sheets("Glastonbury-Regression").Cells(numrows, 9).Value = 1
Else
Sheets("Glastonbury-Regression").Cells(numrows, 9).Value = 0
End If

If newhum = "Dry" Then
Sheets("Glastonbury-Regression").Cells(numrows, 10).Value = 1
Else
Sheets("Glastonbury-Regression").Cells(numrows, 10).Value = 0
End If

If newfest = "Yes" Then
Sheets("Glastonbury-Regression").Cells(numrows, 11).Value = 1
Else
Sheets("Glastonbury-Regression").Cells(numrows, 11).Value = 0
End If
Sheets("Glastonbury-Regression").Cells(numrows, 12).Value = newglaston

'Hartford
Sheets("Hartford - Regression").Cells(numrows, 2).Value = newdate
Sheets("Hartford - Regression").Cells(numrows, 3).Value = newday
Sheets("Hartford - Regression").Cells(numrows, 4).Value = newtemp
Sheets("Hartford - Regression").Cells(numrows, 5).Value = newhum
Sheets("Hartford - Regression").Cells(numrows, 6).Value = newfest
Sheets("Hartford - Regression").Cells(numrows, 7).Value = Sheets("Hartford - Regression").Cells(numrows - 1, 7).Value + 1
If newtemp = "Mild" Then
Sheets("Hartford - Regression").Cells(numrows, 8).Value = 1
Else
Sheets("Hartford - Regression").Cells(numrows, 8).Value = 0
End If
If newtemp = "Cool" Then
Sheets("Hartford - Regression").Cells(numrows, 9).Value = 1
Else
Sheets("Hartford - Regression").Cells(numrows, 9).Value = 0
End If
If newhum = "Dry" Then
Sheets("Hartford - Regression").Cells(numrows, 10).Value = 1
Else
Sheets("Hartford - Regression").Cells(numrows, 10).Value = 0
End If
If newhum = "Very Dry" Then
Sheets("Hartford - Regression").Cells(numrows, 11).Value = 1
Else
Sheets("Hartford - Regression").Cells(numrows, 11).Value = 0
End If

If newfest = "Yes" Then
Sheets("Hartford - Regression").Cells(numrows, 12).Value = 1
Else
Sheets("Hartford - Regression").Cells(numrows, 12).Value = 0
End If
Sheets("Hartford - Regression").Cells(numrows, 13).Value = newhartford

'Manchester
Sheets("Manchester-Regression").Cells(numrows, 2).Value = newdate
Sheets("Manchester-Regression").Cells(numrows, 3).Value = newday
Sheets("Manchester-Regression").Cells(numrows, 4).Value = newtemp
Sheets("Manchester-Regression").Cells(numrows, 5).Value = newhum
Sheets("Manchester-Regression").Cells(numrows, 6).Value = newfest
Sheets("Manchester-Regression").Cells(numrows, 7).Value = Sheets("Manchester-Regression").Cells(numrows - 1, 7).Value + 1
If newfest = "Yes" Then
Sheets("Manchester-Regression").Cells(numrows, 8).Value = 1
Else
Sheets("Manchester-Regression").Cells(numrows, 8).Value = 0
End If
Sheets("Manchester-Regression").Cells(numrows, 9).Value = newhartford

'Southington

Sheets("Southington-Regression").Cells(numrows, 2).Value = newdate
Sheets("Southington-Regression").Cells(numrows, 3).Value = newday
Sheets("Southington-Regression").Cells(numrows, 4).Value = newtemp
Sheets("Southington-Regression").Cells(numrows, 5).Value = newhum
Sheets("Southington-Regression").Cells(numrows, 6).Value = newfest
Sheets("Southington-Regression").Cells(numrows, 7).Value = Sheets("Southington-Regression").Cells(numrows - 1, 7).Value + 1
If newtemp = "Mild" Then
Sheets("Southington-Regression").Cells(numrows, 8).Value = 1
Else
Sheets("Southington-Regression").Cells(numrows, 8).Value = 0
End If
If newtemp = "Warm" Then
Sheets("Southington-Regression").Cells(numrows, 9).Value = 1
Else
Sheets("Southington-Regression").Cells(numrows, 9).Value = 0
End If
If newhum = "Dry" Then
Sheets("Southington-Regression").Cells(numrows, 10).Value = 1
Else
Sheets("Southington-Regression").Cells(numrows, 10).Value = 0
End If
If newhum = "Very Dry" Then
Sheets("Southington-Regression").Cells(numrows, 11).Value = 1
Else
Sheets("Southington-Regression").Cells(numrows, 11).Value = 0
End If

If newfest = "Yes" Then
Sheets("Southington-Regression").Cells(numrows, 12).Value = 1
Else
Sheets("Southington-Regression").Cells(numrows, 12).Value = 0
End If
Sheets("Southington-Regression").Cells(numrows, 13).Value = newsouthington

'West Hartford
Sheets("West Hartford-Regression").Cells(numrows, 2).Value = newdate
Sheets("West Hartford-Regression").Cells(numrows, 3).Value = newday
Sheets("West Hartford-Regression").Cells(numrows, 4).Value = newtemp
Sheets("West Hartford-Regression").Cells(numrows, 5).Value = newhum
Sheets("West Hartford-Regression").Cells(numrows, 6).Value = newfest
Sheets("West Hartford-Regression").Cells(numrows, 7).Value = Sheets("West Hartford-Regression").Cells(numrows - 1, 7).Value + 1
If newtemp = "Mild" Then
Sheets("West Hartford-Regression").Cells(numrows, 8).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 8).Value = 0
End If
If newtemp = "Cool" Then
Sheets("West Hartford-Regression").Cells(numrows, 9).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 9).Value = 0
End If
If newhum = "Dry" Then
Sheets("West Hartford-Regression").Cells(numrows, 10).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 10).Value = 0
End If
If newhum = "Humid" Then
Sheets("West Hartford-Regression").Cells(numrows, 11).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 11).Value = 0
End If
If newhum = "Very Dry" Then
Sheets("West Hartford-Regression").Cells(numrows, 12).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 12).Value = 0
End If

If newfest = "Yes" Then
Sheets("West Hartford-Regression").Cells(numrows, 13).Value = 1
Else
Sheets("West Hartford-Regression").Cells(numrows, 13).Value = 0
End If
Sheets("West Hartford-Regression").Cells(numrows, 14).Value = newwesthartford

'Windsor-Regression
Sheets("Windsor-Regression").Cells(numrows, 2).Value = newdate
Sheets("Windsor-Regression").Cells(numrows, 3).Value = newday
Sheets("Windsor-Regression").Cells(numrows, 4).Value = newtemp
Sheets("Windsor-Regression").Cells(numrows, 5).Value = newhum
Sheets("Windsor-Regression").Cells(numrows, 6).Value = newfest
Sheets("Windsor-Regression").Cells(numrows, 7).Value = Sheets("Windsor-Regression").Cells(numrows - 1, 7).Value + 1

If newhum = "Dry" Then
Sheets("Windsor-Regression").Cells(numrows, 8).Value = 1
Else
Sheets("Windsor-Regression").Cells(numrows, 8).Value = 0
End If
If newhum = "Humid" Then
Sheets("Windsor-Regression").Cells(numrows, 9).Value = 1
Else
Sheets("Windsor-Regression").Cells(numrows, 9).Value = 0
End If
If newhum = "Very dry" Then
Sheets("Windsor-Regression").Cells(numrows, 10).Value = 1
Else
Sheets("Windsor-Regression").Cells(numrows, 10).Value = 0
End If

If newfest = "Yes" Then
Sheets("Windsor-Regression").Cells(numrows, 11).Value = 1
Else
Sheets("Windsor-Regression").Cells(numrows, 11).Value = 0
End If
Sheets("Windsor-Regression").Cells(numrows, 12).Value = newwindsor

numrows = numrows + 1

Next I

'Bloomfield
MsgBox "Performing regression for BLOOMFIELD"
Set sheetname = ThisWorkbook.Sheets("Bloomfield-Regression")
Call Regress(sheetname.Range("J3:J" & totalrows), sheetname.Range("G3:H" & totalrows), False, True, False, sheetname.Range("U3"), False, False, False, False, False, False, False)
Call CalBloomfield(sheetname, totalrows)

'Farmington
MsgBox "Performing regression for FARMINGTON"
Set sheetname = ThisWorkbook.Sheets("Farmington-Regression")
Call Regress(sheetname.Range("M3:M" & totalrows), sheetname.Range("G3:L" & totalrows), False, True, False, sheetname.Range("X30"), False, False, False, False, False, False, False)
Call CalFarmington(sheetname, totalrows)

'Glastonbury
MsgBox "Performing regression for GLASTONBURY"
Set sheetname = ThisWorkbook.Sheets("Glastonbury-Regression")
Call Regress(sheetname.Range("L3:L" & totalrows), sheetname.Range("G3:K" & totalrows), False, True, False, sheetname.Range("W29"), False, False, False, False, False, False, False)
Call CalGlastonbury(sheetname, totalrows)


'Hartford
MsgBox "Performing regression for HARTFORD"
Set sheetname = ThisWorkbook.Sheets("Hartford - Regression")
Call Regress(sheetname.Range("M3:M" & totalrows), sheetname.Range("G3:L" & totalrows), False, True, False, sheetname.Range("X30"), False, False, False, False, False, False, False)
Call CalHartford(sheetname, totalrows)

'Manchester
MsgBox "Performing regression for MANCHESTER"
Set sheetname = ThisWorkbook.Sheets("Manchester-Regression")
Call Regress(sheetname.Range("I3:I" & totalrows), sheetname.Range("G3:H" & totalrows), False, True, False, sheetname.Range("T3"), False, False, False, False, False, False, False)
Call CalManchester(sheetname, totalrows)

'Southington-Regression
MsgBox "Performing regression for SOUTHINGTON"
Set sheetname = ThisWorkbook.Sheets("Southington-Regression")
Call Regress(sheetname.Range("M3:M" & totalrows), sheetname.Range("G3:L" & totalrows), False, True, False, sheetname.Range("X29"), False, False, False, False, False, False, False)
Call CalSouthington(sheetname, totalrows)

'West Hartford -Regression
MsgBox "Performing regression for WEST HARTFORD"
Set sheetname = ThisWorkbook.Sheets("West Hartford-Regression")
Call Regress(sheetname.Range("N3:N" & totalrows), sheetname.Range("G3:M" & totalrows), False, True, False, sheetname.Range("Y30"), False, False, False, False, False, False, False)
Call CalWestHartford(sheetname, totalrows)

'Windsor - Regression
MsgBox "Performing regression for WINDSOR"
Set sheetname = ThisWorkbook.Sheets("Windsor-Regression")
Call Regress(sheetname.Range("L3:L" & totalrows), sheetname.Range("G3:K" & totalrows), False, True, False, sheetname.Range("W28"), False, False, False, False, False, False, False)
Call CalWindsor(sheetname, totalrows)


Call CopyModule

Unload UserForm2

End Sub



' Calculating the ACTUAL/PREDICTED RATIO

Function Ratio_act_Pred(sheetname, totalrows, act, pred, output)

    numrows = 4
    For J = numrows To totalrows
        actual = sheetname.Range(act & numrows)
        predval = sheetname.Range(pred & numrows)
        actpredratio = actual / predval
        sheetname.Range(output & numrows) = actpredratio
        numrows = numrows + 1
    Next

End Function

' Calculating the Seasonal Factor

Function seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

   
    For J = numrows To maxrows
        seasonalfactor = Application.WorksheetFunction.SumIf(criteria, sheetname.Range(sinput & numrows), sumrange) / _
                         Application.WorksheetFunction.CountIf(criteria, sheetname.Range(sinput & numrows))
        sheetname.Range(output & numrows) = seasonalfactor
        numrows = numrows + 1
    Next

End Function

' Calculating Adjusted Predicted Values


Function adjprediction(sheetname, srange, predrow, output, totalrows)

    numrows = 4
    For J = numrows To totalrows
        adjpred = sheetname.Range(predrow & numrows) * (Application.WorksheetFunction.VLookup(sheetname.Range("C" & numrows), srange, 2, False))
        sheetname.Range(output & numrows) = adjpred
        numrows = numrows + 1
    Next

End Function
Function CalBloomfield(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4
' The below code Calculates the Predicted values with the new Coefficients
For J = numrows To totalrows
    predval = (sheetname.Range("V48") + _
    sheetname.Range("V49") * sheetname.Range("G" & numrows) + _
    sheetname.Range("V50") * sheetname.Range("H" & numrows) + _
    sheetname.Range("V51") * sheetname.Range("I" & numrows))
    sheetname.Range("N" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO -   SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
act = "J"
pred = "N"
output = "O"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   Q4 = SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("O4:O" & totalrows)
sinput = "R"
output = "S"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES =L4*VLOOKUP(C4,$P$4:$Q$10,2,FALSE)
Set srange = sheetname.Range("R19:S25")
predrow = "N"
output = "P"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("J4:J" & totalrows), sheetname.Range("N4:N" & totalrows))) / (totalrows - 3)
sheetname.Range("S27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("J4:J" & totalrows), sheetname.Range("P4:P" & totalrows))) / (totalrows - 3)
sheetname.Range("S28") = mserevadj

End Function
Function CalFarmington(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4
' The below code Calculates the Predicted values with the new Coefficients
For J = numrows To totalrows
    predval = (sheetname.Range("Y19") + _
    sheetname.Range("Y20") * sheetname.Range("G" & numrows) + _
    sheetname.Range("Y21") * sheetname.Range("H" & numrows) + _
    sheetname.Range("Y22") * sheetname.Range("I" & numrows) + _
    sheetname.Range("Y23") * sheetname.Range("J" & numrows) + _
    sheetname.Range("Y24") * sheetname.Range("K" & numrows) + _
    sheetname.Range("Y25") * sheetname.Range("L" & numrows))
    sheetname.Range("Q" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO -   SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
act = "M"
pred = "Q"
output = "R"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   Q4 = SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("R4:R" & totalrows)
sinput = "U"
output = "V"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES =L4*VLOOKUP(C4,$P$4:$Q$10,2,FALSE)
Set srange = sheetname.Range("U19:V25")
predrow = "Q"
output = "S"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("Q4:Q" & totalrows))) / (totalrows - 3)
sheetname.Range("V27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("S4:S" & totalrows))) / (totalrows - 3)
sheetname.Range("V28") = mserevadj
'
End Function
Function CalGlastonbury(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4
' The below code Calculates the Predicted values with the new Coefficients
For J = numrows To totalrows
    predval = (sheetname.Range("X45") + _
    sheetname.Range("X46") * sheetname.Range("G" & numrows) + _
    sheetname.Range("X46") * sheetname.Range("H" & numrows) + _
    sheetname.Range("X46") * sheetname.Range("I" & numrows) + _
    sheetname.Range("X46") * sheetname.Range("J" & numrows) + _
    sheetname.Range("X46") * sheetname.Range("K" & numrows))
    sheetname.Range("P" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO -   SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
act = "L"
pred = "P"
output = "Q"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   Q4 = SUMIF($C$4:$C$128,P4,$M$4:$M$128)/COUNTIF($C$4:$C$128,P4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("Q4:Q" & totalrows)
sinput = "T"
output = "U"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES =L4*VLOOKUP(C4,$P$4:$Q$10,2,FALSE)
Set srange = sheetname.Range("T19:U25")
predrow = "P"
output = "R"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("L4:L" & totalrows), sheetname.Range("P4:P" & totalrows))) / (totalrows - 3)
sheetname.Range("U27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("L4:L" & totalrows), sheetname.Range("R4:R" & totalrows))) / (totalrows - 3)
sheetname.Range("U28") = mserevadj
 
End Function
Function CalHartford(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4

'PREDICTED VALUES

For J = numrows To totalrows
    predval = (sheetname.Range("Y46") + _
    sheetname.Range("Y47") * sheetname.Range("G" & numrows) + _
    sheetname.Range("Y48") * sheetname.Range("H" & numrows) + _
    sheetname.Range("Y49") * sheetname.Range("I" & numrows) + _
    sheetname.Range("Y50") * sheetname.Range("J" & numrows) + _
    sheetname.Range("Y51") * sheetname.Range("K" & numrows) + _
    sheetname.Range("Y52") * sheetname.Range("L" & numrows))
    sheetname.Range("Q" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO
act = "M"
pred = "Q"
output = "R"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   O4 = SUMIF($C$4:$C$128,N4,$K$4:$K$128)/COUNTIF($C$4:$C$128,N4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("R4:R" & totalrows)
sinput = "U"
output = "V"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES  L4 = J4*VLOOKUP(C4,$N$4:$O$10,2,FALSE)
Set srange = sheetname.Range("U19:V25")
predrow = "Q"
output = "S"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("Q4:Q" & totalrows))) / (totalrows - 3)
sheetname.Range("V27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("S4:S" & totalrows))) / (totalrows - 3)
sheetname.Range("V28") = mserevadj


End Function


Function CalManchester(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4

'PREDICTED VALUES

For J = numrows To totalrows
    predval = (sheetname.Range("U44") + _
    sheetname.Range("U45") * sheetname.Range("G" & numrows) + _
    sheetname.Range("U46") * sheetname.Range("H" & numrows))
    sheetname.Range("M" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO
act = "I"
pred = "M"
output = "N"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   O4 = SUMIF($C$4:$C$128,N4,$K$4:$K$128)/COUNTIF($C$4:$C$128,N4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("N4:N" & totalrows)
sinput = "Q"
output = "R"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES  L4 = J4*VLOOKUP(C4,$N$4:$O$10,2,FALSE)
Set srange = sheetname.Range("Q19:R25")
predrow = "M"
output = "O"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("I4:I" & totalrows), sheetname.Range("M4:M" & totalrows))) / (totalrows - 3)
sheetname.Range("R27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("I4:I" & totalrows), sheetname.Range("O4:O" & totalrows))) / (totalrows - 3)
sheetname.Range("R28") = mserevadj


End Function
Function CalSouthington(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4

'PREDICTED VALUES

For J = numrows To totalrows
    predval = (sheetname.Range("Y45") + _
    sheetname.Range("Y46") * sheetname.Range("G" & numrows) + _
    sheetname.Range("Y47") * sheetname.Range("H" & numrows) + _
    sheetname.Range("Y48") * sheetname.Range("I" & numrows) + _
    sheetname.Range("Y49") * sheetname.Range("J" & numrows) + _
    sheetname.Range("Y50") * sheetname.Range("K" & numrows) + _
    sheetname.Range("Y51") * sheetname.Range("L" & numrows))
    sheetname.Range("Q" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO
act = "M"
pred = "Q"
output = "R"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  -   S4 = SUMIF($C$4:$C$128,R4,$O$4:$O$128)/COUNTIF($C$4:$C$128,R4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("R4:R" & totalrows)
numrows = 19
maxrows = 25
sinput = "U"
output = "V"
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)

'ADJUSTED PREDICTION VALUES  P4 = N4*VLOOKUP(C4,$R$4:$S$10,2,FALSE)
Set srange = sheetname.Range("U19:V25")
predrow = "Q"
output = "S"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("Q4:Q" & totalrows))) / (totalrows - 3)
sheetname.Range("V27") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("M4:M" & totalrows), sheetname.Range("S4:S" & totalrows))) / (totalrows - 3)
sheetname.Range("V28") = mserevadj

End Function

Function CalWestHartford(sheetname, totalrows)
 
Set sheetname = sheetname
totalrows = totalrows

numrows = 4

'PREDICTED VALUES
For J = numrows To totalrows
    predval = (sheetname.Range("Z46") + _
    sheetname.Range("Z47") * sheetname.Range("G" & numrows) + _
    sheetname.Range("Z48") * sheetname.Range("H" & numrows) + _
    sheetname.Range("Z49") * sheetname.Range("I" & numrows) + _
    sheetname.Range("Z50") * sheetname.Range("J" & numrows) + _
    sheetname.Range("Z51") * sheetname.Range("K" & numrows) + _
    sheetname.Range("Z52") * sheetname.Range("L" & numrows) + _
    sheetname.Range("Z53") * sheetname.Range("M" & numrows))
    sheetname.Range("R" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO
act = "N"
pred = "R"
output = "S"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  - T4 = SUMIF($C$4:$C$128,S4,$P$4:$P$128)/COUNTIF($C$4:$C$128,S4)
Set criteria = sheetname.Range("$C$4:$C$" & totalrows)
Set sumrange = sheetname.Range("$S$4:$S$" & totalrows)
 
sinput = "V"
output = "W"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)


'ADJUSTED PREDICTION VALUES  Q4 =N4*VLOOKUP(C4,$S$4:$T$10,2,FALSE)
Set srange = sheetname.Range("V19:W25")
predrow = "R"
output = "T"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mserev = (Application.WorksheetFunction.SumXMY2(sheetname.Range("N4:N" & totalrows), sheetname.Range("R4:R" & totalrows))) / (totalrows - 3)
sheetname.Range("W15") = mserev

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("N4:N" & totalrows), sheetname.Range("T4:T" & totalrows))) / (totalrows - 3)
sheetname.Range("W16") = mserevadj

End Function
Function CalWindsor(sheetname, totalrows)

Set sheetname = sheetname
totalrows = totalrows

numrows = 4
' The below code Calculates the Predicted values with the new Coefficients
' P4
For J = numrows To totalrows
    predval = (sheetname.Range("X44") + _
    sheetname.Range("X45") * sheetname.Range("G" & numrows) + _
    sheetname.Range("X46") * sheetname.Range("H" & numrows) + _
    sheetname.Range("X47") * sheetname.Range("I" & numrows) + _
    sheetname.Range("X48") * sheetname.Range("J" & numrows) + _
    sheetname.Range("X49") * sheetname.Range("K" & numrows))
    sheetname.Range("P" & numrows) = predval
    numrows = numrows + 1
Next

'ACTUAL/PREDICTED RAIO -   Q4 = L4/P4
act = "L"
pred = "P"
output = "Q"
Call Ratio_act_Pred(sheetname, totalrows, act, pred, output)

'SEASONAL FACTOR  - U19 = SUMIF($C$4:$C$128,T4,$Q$4:$Q$128)/COUNTIF($C$4:$C$128,T4)
Set criteria = sheetname.Range("C4:C" & totalrows)
Set sumrange = sheetname.Range("Q4:Q" & totalrows)
sinput = "T"
output = "U"
numrows = 19
maxrows = 25
Call seasonal(sheetname, criteria, sumrange, sinput, output, numrows, maxrows)


'ADJUSTED PREDICTION VALUES  R4 =P4*VLOOKUP(C4,$T$19:$U$25,2,FALSE)
Set srange = sheetname.Range("T19:U25")
predrow = "P"
output = "R"
Call adjprediction(sheetname, srange, predrow, output, totalrows)

mse = (Application.WorksheetFunction.SumXMY2(sheetname.Range("L4:L" & totalrows), sheetname.Range("P4:P" & totalrows))) / (totalrows - 3)
sheetname.Range("U28") = mse

mserevadj = (Application.WorksheetFunction.SumXMY2(sheetname.Range("L4:L" & totalrows), sheetname.Range("R4:R" & totalrows))) / (totalrows - 3)
sheetname.Range("U29") = mserevadj

End Function
Function CopyModule()

'Clears the contents of the cells
Sheets("Model Summary").Range("B3:U10").Clear
 
'Bloomfield
' Revised Coefficients
Sheets("Model Summary").Range("B3") = Sheets("Bloomfield-Regression").Range("V48")
Sheets("Model Summary").Range("C3") = Sheets("Bloomfield-Regression").Range("V49")
Sheets("Model Summary").Range("E3") = Sheets("Bloomfield-Regression").Range("V50")
Sheets("Model Summary").Range("K3") = Sheets("Bloomfield-Regression").Range("V51")
'F-Statistic
Sheets("Model Summary").Range("L3") = Sheets("Bloomfield-Regression").Range("Z43")
'Seasonal Coefficients
Sheets("Model Summary").Range("M3") = Sheets("Bloomfield-Regression").Range("S19")
Sheets("Model Summary").Range("N3") = Sheets("Bloomfield-Regression").Range("S20")
Sheets("Model Summary").Range("O3") = Sheets("Bloomfield-Regression").Range("S21")
Sheets("Model Summary").Range("P3") = Sheets("Bloomfield-Regression").Range("S22")
Sheets("Model Summary").Range("Q3") = Sheets("Bloomfield-Regression").Range("S23")
Sheets("Model Summary").Range("R3") = Sheets("Bloomfield-Regression").Range("S24")
Sheets("Model Summary").Range("S3") = Sheets("Bloomfield-Regression").Range("S25")
'MSE
Sheets("Model Summary").Range("T3") = Sheets("Bloomfield-Regression").Range("S27")
Sheets("Model Summary").Range("U3") = Sheets("Bloomfield-Regression").Range("S28")


'Farmington
' Revised Coefficients
Sheets("Model Summary").Range("B4") = Sheets("Farmington-Regression").Range("Y46")
Sheets("Model Summary").Range("C4") = Sheets("Farmington-Regression").Range("Y47")
Sheets("Model Summary").Range("D4") = Sheets("Farmington-Regression").Range("Y48")
Sheets("Model Summary").Range("E4") = Sheets("Farmington-Regression").Range("Y49")
Sheets("Model Summary").Range("H4") = Sheets("Farmington-Regression").Range("Y50")
Sheets("Model Summary").Range("I4") = Sheets("Farmington-Regression").Range("Y51")
Sheets("Model Summary").Range("K4") = Sheets("Farmington-Regression").Range("Y52")
'F-Statistic
Sheets("Model Summary").Range("L4") = Sheets("Farmington-Regression").Range("AC41")
'Seasonal Coefficients
Sheets("Model Summary").Range("M4") = Sheets("Farmington-Regression").Range("V19")
Sheets("Model Summary").Range("N4") = Sheets("Farmington-Regression").Range("V20")
Sheets("Model Summary").Range("O4") = Sheets("Farmington-Regression").Range("V21")
Sheets("Model Summary").Range("P4") = Sheets("Farmington-Regression").Range("V22")
Sheets("Model Summary").Range("Q4") = Sheets("Farmington-Regression").Range("V23")
Sheets("Model Summary").Range("R4") = Sheets("Farmington-Regression").Range("V24")
Sheets("Model Summary").Range("S4") = Sheets("Farmington-Regression").Range("V25")
'MSE
Sheets("Model Summary").Range("T4") = Sheets("Farmington-Regression").Range("V27")
Sheets("Model Summary").Range("U4") = Sheets("Farmington-Regression").Range("V28")


'Glastonbury

'Revised Coefficients
Sheets("Model Summary").Range("B5") = Sheets("Glastonbury-Regression").Range("X45")
Sheets("Model Summary").Range("C5") = Sheets("Glastonbury-Regression").Range("X46")
Sheets("Model Summary").Range("E5") = Sheets("Glastonbury-Regression").Range("X47")
Sheets("Model Summary").Range("F5") = Sheets("Glastonbury-Regression").Range("X48")
Sheets("Model Summary").Range("H5") = Sheets("Glastonbury-Regression").Range("X49")
Sheets("Model Summary").Range("K5") = Sheets("Glastonbury-Regression").Range("X50")

'F-Statistic
Sheets("Model Summary").Range("L5") = Sheets("Glastonbury-Regression").Range("AB40")
'Seasonal Coefficients
Sheets("Model Summary").Range("M5") = Sheets("Glastonbury-Regression").Range("U19")
Sheets("Model Summary").Range("N5") = Sheets("Glastonbury-Regression").Range("U20")
Sheets("Model Summary").Range("O5") = Sheets("Glastonbury-Regression").Range("U21")
Sheets("Model Summary").Range("P5") = Sheets("Glastonbury-Regression").Range("U22")
Sheets("Model Summary").Range("Q5") = Sheets("Glastonbury-Regression").Range("U23")
Sheets("Model Summary").Range("R5") = Sheets("Glastonbury-Regression").Range("U24")
Sheets("Model Summary").Range("S5") = Sheets("Glastonbury-Regression").Range("U25")

'MSE
Sheets("Model Summary").Range("T5") = Sheets("Glastonbury-Regression").Range("U27")
Sheets("Model Summary").Range("U5") = Sheets("Glastonbury-Regression").Range("U28")


'Hartford
'Revised Coefficients
Sheets("Model Summary").Range("B6") = Sheets("Hartford - Regression").Range("Y46")
Sheets("Model Summary").Range("C6") = Sheets("Hartford - Regression").Range("Y47")
Sheets("Model Summary").Range("D6") = Sheets("Hartford - Regression").Range("Y48")
Sheets("Model Summary").Range("F6") = Sheets("Hartford - Regression").Range("Y49")
Sheets("Model Summary").Range("H6") = Sheets("Hartford - Regression").Range("Y50")
Sheets("Model Summary").Range("I6") = Sheets("Hartford - Regression").Range("Y51")
Sheets("Model Summary").Range("K6") = Sheets("Hartford - Regression").Range("Y52")

'F-Statistic
Sheets("Model Summary").Range("L6") = Sheets("Hartford - Regression").Range("AC41")
'Seasonal Coefficients
Sheets("Model Summary").Range("M6") = Sheets("Hartford - Regression").Range("V19")
Sheets("Model Summary").Range("N6") = Sheets("Hartford - Regression").Range("V20")
Sheets("Model Summary").Range("O6") = Sheets("Hartford - Regression").Range("V21")
Sheets("Model Summary").Range("P6") = Sheets("Hartford - Regression").Range("V22")
Sheets("Model Summary").Range("Q6") = Sheets("Hartford - Regression").Range("V23")
Sheets("Model Summary").Range("R6") = Sheets("Hartford - Regression").Range("V24")
Sheets("Model Summary").Range("S6") = Sheets("Hartford - Regression").Range("V25")

'MSE
Sheets("Model Summary").Range("T6") = Sheets("Hartford - Regression").Range("V27")
Sheets("Model Summary").Range("U6") = Sheets("Hartford - Regression").Range("V28")


'Manchester-Regression
'Revised Coefficients
Sheets("Model Summary").Range("B7") = Sheets("Manchester-Regression").Range("U44")
Sheets("Model Summary").Range("C7") = Sheets("Manchester-Regression").Range("U45")
Sheets("Model Summary").Range("K7") = Sheets("Manchester-Regression").Range("U46")

'F-Stat
Sheets("Model Summary").Range("L7") = Sheets("Manchester-Regression").Range("Y39")

'Seasonal Coefficients
Sheets("Model Summary").Range("M7") = Sheets("Manchester-Regression").Range("R19")
Sheets("Model Summary").Range("N7") = Sheets("Manchester-Regression").Range("R20")
Sheets("Model Summary").Range("O7") = Sheets("Manchester-Regression").Range("R21")
Sheets("Model Summary").Range("P7") = Sheets("Manchester-Regression").Range("R22")
Sheets("Model Summary").Range("Q7") = Sheets("Manchester-Regression").Range("R23")
Sheets("Model Summary").Range("R7") = Sheets("Manchester-Regression").Range("R24")
Sheets("Model Summary").Range("S7") = Sheets("Manchester-Regression").Range("R25")

'MSE
Sheets("Model Summary").Range("T7") = Sheets("Manchester-Regression").Range("R27")
Sheets("Model Summary").Range("U7") = Sheets("Manchester-Regression").Range("R28")

'Southington-Regression
Sheets("Model Summary").Range("B8") = Sheets("Southington-Regression").Range("Y45")
Sheets("Model Summary").Range("C8") = Sheets("Southington-Regression").Range("Y46")
Sheets("Model Summary").Range("D8") = Sheets("Southington-Regression").Range("Y47")
Sheets("Model Summary").Range("E8") = Sheets("Southington-Regression").Range("Y48")
Sheets("Model Summary").Range("H8") = Sheets("Southington-Regression").Range("Y49")
Sheets("Model Summary").Range("I8") = Sheets("Southington-Regression").Range("Y50")
Sheets("Model Summary").Range("K8") = Sheets("Southington-Regression").Range("Y51")

'F-Stat
Sheets("Model Summary").Range("L8") = Sheets("Southington-Regression").Range("AC40")

'Seasonal
Sheets("Model Summary").Range("M8") = Sheets("Southington-Regression").Range("V19")
Sheets("Model Summary").Range("N8") = Sheets("Southington-Regression").Range("V20")
Sheets("Model Summary").Range("O8") = Sheets("Southington-Regression").Range("V21")
Sheets("Model Summary").Range("P8") = Sheets("Southington-Regression").Range("V22")
Sheets("Model Summary").Range("Q8") = Sheets("Southington-Regression").Range("V23")
Sheets("Model Summary").Range("R8") = Sheets("Southington-Regression").Range("V24")
Sheets("Model Summary").Range("S8") = Sheets("Southington-Regression").Range("V25")
'MSE
Sheets("Model Summary").Range("T8") = Sheets("Southington-Regression").Range("V27")
Sheets("Model Summary").Range("U8") = Sheets("Southington-Regression").Range("V28")

'West Hartford-Regression

Sheets("Model Summary").Range("B9") = Sheets("West Hartford-Regression").Range("Z46")
Sheets("Model Summary").Range("C9") = Sheets("West Hartford-Regression").Range("Z47")
Sheets("Model Summary").Range("D9") = Sheets("West Hartford-Regression").Range("Z48")
Sheets("Model Summary").Range("F9") = Sheets("West Hartford-Regression").Range("Z49")
Sheets("Model Summary").Range("H9") = Sheets("West Hartford-Regression").Range("Z50")
Sheets("Model Summary").Range("J9") = Sheets("West Hartford-Regression").Range("Z51")
Sheets("Model Summary").Range("I9") = Sheets("West Hartford-Regression").Range("Z52")
Sheets("Model Summary").Range("K9") = Sheets("West Hartford-Regression").Range("Z53")

'F-Stat
Sheets("Model Summary").Range("L9") = Sheets("West Hartford-Regression").Range("AD41")

'Seasonal
Sheets("Model Summary").Range("M9") = Sheets("West Hartford-Regression").Range("W19")
Sheets("Model Summary").Range("N9") = Sheets("West Hartford-Regression").Range("W20")
Sheets("Model Summary").Range("O9") = Sheets("West Hartford-Regression").Range("W21")
Sheets("Model Summary").Range("P9") = Sheets("West Hartford-Regression").Range("W22")
Sheets("Model Summary").Range("Q9") = Sheets("West Hartford-Regression").Range("W23")
Sheets("Model Summary").Range("R9") = Sheets("West Hartford-Regression").Range("W24")
Sheets("Model Summary").Range("S9") = Sheets("West Hartford-Regression").Range("W25")

'MSE
Sheets("Model Summary").Range("T9") = Sheets("West Hartford-Regression").Range("W15")
Sheets("Model Summary").Range("U9") = Sheets("West Hartford-Regression").Range("W16")

'Windsor-Regression
Sheets("Model Summary").Range("B10") = Sheets("Windsor-Regression").Range("X44")
Sheets("Model Summary").Range("C10") = Sheets("Windsor-Regression").Range("X45")
Sheets("Model Summary").Range("H10") = Sheets("Windsor-Regression").Range("X46")
Sheets("Model Summary").Range("J10") = Sheets("Windsor-Regression").Range("X47")
Sheets("Model Summary").Range("I10") = Sheets("Windsor-Regression").Range("X48")
Sheets("Model Summary").Range("K10") = Sheets("Windsor-Regression").Range("X49")
'F -STAT
Sheets("Model Summary").Range("L10") = Sheets("Windsor-Regression").Range("AB39")
'Seasonal
Sheets("Model Summary").Range("M10") = Sheets("Windsor-Regression").Range("U19")
Sheets("Model Summary").Range("N10") = Sheets("Windsor-Regression").Range("U20")
Sheets("Model Summary").Range("O10") = Sheets("Windsor-Regression").Range("U21")
Sheets("Model Summary").Range("P10") = Sheets("Windsor-Regression").Range("U22")
Sheets("Model Summary").Range("Q10") = Sheets("Windsor-Regression").Range("U23")
Sheets("Model Summary").Range("R10") = Sheets("Windsor-Regression").Range("U24")
Sheets("Model Summary").Range("S10") = Sheets("Windsor-Regression").Range("U25")
'MSE
Sheets("Model Summary").Range("T10") = Sheets("Windsor-Regression").Range("U28")
Sheets("Model Summary").Range("U10") = Sheets("Windsor-Regression").Range("U29")

End Function

Private Sub UserForm_Initialize()
Dim recentdate As Date
ComboBox1.AddItem "Cold"
ComboBox1.AddItem "Cool"
ComboBox1.AddItem "Mild"
ComboBox1.AddItem "Warm"
ComboBox1.AddItem "Hot"
ComboBox1.AddItem "Very Hot"
ComboBox2.AddItem "Dry"
ComboBox2.AddItem "Humid"
ComboBox2.AddItem "Very Dry"
ComboBox2.AddItem "Very Humid"
ListBox1.ColumnCount = 4
ListBox1.AddItem
cnt = ListBox1.ListCount - 1
ListBox1.List(cnt, 0) = "Date"
ListBox1.List(cnt, 1) = "Temperature"
ListBox1.List(cnt, 2) = "Humidity"
ListBox1.List(cnt, 3) = "Festive Day"

J = Sheets("Orders Summary").Range("B3").End(xlDown).Row
recentdate = Sheets("Orders Summary").Cells(J, "B")

recentdate = DateAdd("d", 1, recentdate)
Label5.Caption = recentdate

End Sub
