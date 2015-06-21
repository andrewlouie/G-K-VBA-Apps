VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents TBGroup As MSForms.TextBox
Attribute TBGroup.VB_VarHelpID = -1
Public Function SafeVlookup(lookup_value, table_array, _
                        col_index, range_lookup, error_value) As Variant
    On Error Resume Next
    Err.Clear
    return_value = Application.WorksheetFunction.VLookup(lookup_value, _
                                table_array, col_index, range_lookup)
    If Err <> 0 Then
      return_value = error_value
    End If

    SafeVlookup = return_value
    On Error GoTo 0
End Function
Private Sub TBGroup_Change()
    If loading <> 1 And scrolling <> 1 And loading2 <> 1 Then
    loading2 = 1
    aa = 1
For aa = 1 To 8
For bb = 0 To 10
arrays(bb, aa + UserForm1.CheckBox1.Caption - 1) = UserForm1.Controls("Item" & aa & bb).Text
Next
arrays(11, aa + UserForm1.CheckBox1.Caption - 1) = UserForm1.Controls("Checkbox" & aa).Value
If SafeVlookup(Left(arrays(2, aa + UserForm1.CheckBox1.Caption - 1), 8), ThisWorkbook.Worksheets("Sheet2").Range("H:I"), 2, False, 0) > 0 Or (arrays(10, aa + UserForm1.CheckBox1.Caption - 1) = "OS" And arrays(2, aa + UserForm1.CheckBox1.Caption - 1) = arrays(2, aa + UserForm1.CheckBox1.Caption - 2)) Then UserForm1.Controls("Info" & aa).Visible = True Else UserForm1.Controls("info" & aa).Visible = False
Next
Total1 = 0
Total2 = 0
Erase image, prep
aa = 1
For aa = 1 To 8
UserForm1.Controls("Item" & aa & 5).Text = Round(Val(UserForm1.Controls("Item" & aa & 3).Text) * Val(UserForm1.Controls("Item" & aa & 4).Text), 2)
UserForm1.Controls("Item" & aa & 7).Text = Round(Val(UserForm1.Controls("Item" & aa & 3).Text) * Val(UserForm1.Controls("Item" & aa & 6).Text), 2)
UserForm1.Controls("Item" & aa & 8).Text = Round(UserForm1.Controls("Item" & aa & 5).Text - UserForm1.Controls("Item" & aa & 7).Text, 2)
Next
For aa = 1 To 198
If arrays(11, aa) = True Then
Total1 = Total1 + (Val(arrays(3, aa)) * Val(arrays(4, aa)))
Total2 = Total2 + (Val(arrays(3, aa)) * Val(arrays(6, aa)))
cc = Val(SafeVlookup(Val(Left(arrays(2, aa), 4)), ThisWorkbook.Worksheets("Sheet2").Range("A:C"), 2, False, 0))
dd = Val(SafeVlookup(Val(Left(arrays(2, aa), 4)), ThisWorkbook.Worksheets("Sheet2").Range("A:C"), 3, False, 0))
image(cc) = image(cc) + Val(arrays(0, aa))
prep(dd) = prep(dd) + Val(arrays(0, aa))
here:
End If
Next
If UserForm1.OptionButton1.Value = True Then
If UserForm1.Energy1.Text <> "" Then UserForm1.Energy2.Text = Round(Total1 * (Val(UserForm1.Energy1.Text) / 100), 2) Else UserForm1.Energy2.Text = 0
Else: UserForm1.Energy2.Text = Val(UserForm1.Energy1.Text)
End If
If UserForm1.OptionButton3.Value = True Then
If UserForm1.Energy3.Text <> "" Then UserForm1.Energy4.Text = Round(Total2 * (Val(UserForm1.Energy3.Text) / 100), 2) Else UserForm1.Energy4.Text = 0
Else: UserForm1.Energy4.Text = Val(UserForm1.Energy3.Text)
End If
If UserForm1.OptionButton5.Value = True Then
If UserForm1.Enviro1.Text <> "" Then UserForm1.Enviro2.Text = Round(Total1 * (Val(UserForm1.Enviro1.Text) / 100), 2) Else UserForm1.Enviro2.Text = 0
Else: UserForm1.Enviro2.Text = Val(UserForm1.Enviro1.Text)
End If
If UserForm1.OptionButton7.Value = True Then
If UserForm1.Enviro3.Text <> "" Then UserForm1.Enviro4.Text = Round(Total2 * (Val(UserForm1.Enviro3.Text) / 100), 2) Else UserForm1.Enviro4.Text = 0
Else: UserForm1.Enviro4.Text = Val(UserForm1.Enviro3.Text)
End If
UserForm1.Energy5.Text = Round(Val(UserForm1.Energy2.Text) - Val(UserForm1.Energy4.Text), 2)
UserForm1.Enviro5.Text = Round(Val(UserForm1.Enviro2.Text) - Val(UserForm1.Enviro4.Text), 2)
ee = 1
imageg = 0
prepg = 0
For ee = 1 To 4
If image(ee) > 0 Then imageg = imageg + (SafeVlookup(ee & UserForm1.Image2.Value, ThisWorkbook.Worksheets("Sheet2").Range("D:E"), 2, False, 0) * image(ee))
If prep(ee) > 0 Then prepg = prepg + (SafeVlookup(ee & UserForm1.Prep2.Value, ThisWorkbook.Worksheets("Sheet2").Range("F:G"), 2, False, 0) * prep(ee))
Next
UserForm1.Image3.Text = Round(imageg, 2)
UserForm1.Prep3.Text = Round(prepg, 2)
UserForm1.Image4.Text = Round(Val(UserForm1.Image1.Text) - Val(UserForm1.Image3.Text), 2)
UserForm1.Prep4.Text = Round(Val(UserForm1.Prep1.Text) - Val(UserForm1.Prep3.Text), 2)
UserForm1.Total1.Text = Round(Total1 + IIf(UserForm1.CheckBox11.Value = True, Val(UserForm1.Energy2.Text), 0) + IIf(UserForm1.CheckBox12.Value = True, Val(UserForm1.Enviro2.Text), 0) + IIf(UserForm1.CheckBox13.Value = True, Val(UserForm1.Image1.Text), 0) + IIf(UserForm1.CheckBox14.Value = True, Val(UserForm1.Prep1.Text), 0), 2)
UserForm1.Total2.Text = Round(Total2 + IIf(UserForm1.CheckBox11.Value = True, Val(UserForm1.Energy4.Text), 0) + IIf(UserForm1.CheckBox12.Value = True, Val(UserForm1.Enviro4.Text), 0) + IIf(UserForm1.CheckBox13.Value = True, Val(UserForm1.Image3.Text), 0) + IIf(UserForm1.CheckBox14.Value = True, Val(UserForm1.Prep3.Text), 0), 2)
UserForm1.Total3.Text = Round(UserForm1.Total1.Text - UserForm1.Total2.Text, 2)
loading2 = 0
End If

End Sub
