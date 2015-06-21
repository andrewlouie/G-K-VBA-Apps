VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Quote Maker"
   ClientHeight    =   9285.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11310
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public loaded As Integer
Public total As Integer


'itemform
'0 = item name
'1 = isc
'2 = price
'3 = replacement price
'4 = outsize price
'5 = outsize replacement
'6 = auto replace percentage
'7 = prep
'8 = name tag
'9 = emblem
'10 = colour
'11 = quantity
'12 = frequency
'13 = on/off
'14 = catagory
'15 = original arrays item #
'possible updates:
'default locker numbers
'options for SIC/Competition
'frequency/week for refills to match install date (calender select), don't need abcd for weekly/bi-weekly refills

Private Sub CheckBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub CheckBox9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub ComboBox10_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If scrolling <> 1 And itemform(showing + 1, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox10.Value = itemform(showing + 1, 0)
Exit Sub
End If
OptionButton6.Value = True
Call updateitem(2)
End Sub
Private Sub ComboBox13_Change()
If scrolling <> 1 Then
itemform(showing + 2, 10) = ComboBox13.Value
If ComboBox13.Value <> "" Then TextBox6.Text = Left(TextBox6.Text, 8) & "-" & Left(ComboBox13.Value, 3)
itemform(showing + 2, 1) = TextBox6.Text
End If
End Sub
Private Sub ComboBox14_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If scrolling <> 1 And itemform(showing + 2, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox14.Value = itemform(showing + 2, 0)
Exit Sub
End If
OptionButton8.Value = True
Call updateitem(3)
End Sub
Private Sub ComboBox11_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If loaded = 1 Then
If scrolling <> 1 And itemform(showing + 1, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox11.Value = itemform(showing + 1, 14)
Exit Sub
End If
Call filter(2)
End If
End Sub
Private Sub ComboBox17_Change()
Call calculate
End Sub
Private Sub ComboBox18_Change()
Call calculate
End Sub
Private Sub ComboBox20_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub ComboBox21_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub ComboBox22_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub

Private Sub ComboBox23_Change()
If ComboBox23.ListIndex = 0 Then TextBox75.Text = "" Else TextBox75.Text = refills(ComboBox23.ListIndex - 1, 1)
End Sub
Private Sub ComboBox24_Change()
If ComboBox24.ListIndex = 0 Then TextBox76.Text = "" Else TextBox76.Text = refills(ComboBox24.ListIndex - 1, 1)
End Sub
Private Sub ComboBox25_Change()
If ComboBox25.ListIndex = 0 Then TextBox77.Text = "" Else TextBox77.Text = refills(ComboBox25.ListIndex - 1, 1)
End Sub
Private Sub ComboBox7_Change()
freqchange = 1
Call updateitem(1)
Call ComboBox4_Change
freqchange = 0
End Sub
Private Sub ComboBox8_Change()
freqchange = 1
Call updateitem(2)
Call ComboBox9_Change
freqchange = 0
End Sub
Private Sub ComboBox12_Change()
freqchange = 1
Call updateitem(3)
Call ComboBox13_Change
freqchange = 0
End Sub
Private Sub ComboBox15_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If loaded = 1 Then
If scrolling <> 1 And itemform(showing + 2, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox15.Value = itemform(showing + 2, 14)
Exit Sub
End If
Call filter(3)
End If
End Sub
Private Sub ComboBox16_Change()
If loaded = 1 Then
b = 1
Do While b < lastitem + 1
If itemform(b, 2) = "" Then d = 0 Else d = itemform(b, 2)
If itemform(b, 11) = "" Then e = 0 Else e = itemform(b, 11)
If itemform(b, 13) = True Then C = C + (d * e)
b = b + 1
Loop
If C > 0 Then a = MsgBox("Reset prices?", vbYesNo)
If a = 6 Or C <= 0 Then
b = 1
Do While b < lastitem + 1
If itemform(b, 0) <> "" Then
C = itemform(b, 15)
If ComboBox16.ListIndex = 0 Then
If ComboBox7.ListIndex = 0 Then itemform(b, 2) = arrays(3, C)
If ComboBox7.ListIndex = 1 Then itemform(b, 2) = arrays(4, C)
If ComboBox7.ListIndex = 2 Then itemform(b, 2) = arrays(5, C)
End If
If ComboBox16.ListIndex = 1 Then
If ComboBox7.ListIndex = 0 Then itemform(b, 2) = arrays(6, C)
If ComboBox7.ListIndex = 1 Then itemform(b, 2) = arrays(7, C)
If ComboBox7.ListIndex = 2 Then itemform(b, 2) = arrays(8, C)
End If
If ComboBox16.ListIndex = 2 Then
If ComboBox7.ListIndex = 0 Then itemform(b, 2) = arrays(9, C)
If ComboBox7.ListIndex = 1 Then itemform(b, 2) = arrays(10, C)
If ComboBox7.ListIndex = 2 Then itemform(b, 2) = arrays(11, C)
End If
itemform(showing, 3) = arrays(16, C)
itemform(showing, 4) = Round(itemform(b, 2) + (itemform(b, 2) * arrays(20, C)), 2)
itemform(showing, 5) = Round(arrays(16, C) + (arrays(16, C) * arrays(20, C)), 2)
itemform(showing, 6) = arrays(12, C)
itemform(showing, 7) = arrays(21, C)
itemform(showing, 8) = arrays(22, C)
itemform(showing, 9) = arrays(23, C)
End If
b = b + 1
Loop
TextBox83.Text = Format(itemform(showing, 2), "Currency")
TextBox86.Text = Format(itemform(showing + 1, 2), "Currency")
TextBox89.Text = Format(itemform(showing + 2, 2), "Currency")
Call calculate
End If
End If
End Sub
Private Sub ComboBox4_Change()
If scrolling <> 1 Then
itemform(showing, 10) = ComboBox4.Value
If ComboBox4.Value <> "" Then TextBox3.Text = Left(TextBox3.Text, 8) & "-" & Left(ComboBox4.Value, 3)
itemform(showing, 1) = TextBox3.Text
End If
End Sub
Private Sub ComboBox5_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If scrolling <> 1 And itemform(showing, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox5.Value = itemform(showing, 0)
Exit Sub
End If
OptionButton4.Value = True
Call updateitem(1)
End Sub
Sub updateitem(which As Integer)
If scrolling <> 1 Then
If loaded = 1 Then
If which = 1 And ComboBox5.ListIndex >= 0 Then
C = list1(ComboBox5.ListIndex)
TextBox82.Text = arrays(2, C)
TextBox3.Text = arrays(14, C)
If ComboBox16.ListIndex = 0 Then
If ComboBox7.ListIndex = 0 Then TextBox83.Text = arrays(3, C)
If ComboBox7.ListIndex = 1 Then TextBox83.Text = arrays(4, C)
If ComboBox7.ListIndex = 2 Then TextBox83.Text = arrays(5, C)
End If
If ComboBox16.ListIndex = 1 Then
If ComboBox7.ListIndex = 0 Then TextBox83.Text = arrays(6, C)
If ComboBox7.ListIndex = 1 Then TextBox83.Text = arrays(7, C)
If ComboBox7.ListIndex = 2 Then TextBox83.Text = arrays(8, C)
End If
If ComboBox16.ListIndex = 2 Then
If ComboBox7.ListIndex = 0 Then TextBox83.Text = arrays(9, C)
If ComboBox7.ListIndex = 1 Then TextBox83.Text = arrays(10, C)
If ComboBox7.ListIndex = 2 Then TextBox83.Text = arrays(11, C)
End If
If freqchange <> 1 Then
ComboBox4.Clear
Dim colour() As String
colour = Split(arrays(18, C), ",")
If InStr(arrays(18, C), ",") = 0 Then
ComboBox4.AddItem arrays(18, C)
ComboBox4.ListIndex = 0
TextBox3.Text = Left(TextBox3.Text, 8) & "-" & Left(ComboBox4.Value, 3)
Else
Dim i As Integer
For i = LBound(colour) To UBound(colour)
ComboBox4.AddItem colour(i)
Next i
End If
End If
itemform(showing, 0) = arrays(1, C)
itemform(showing, 1) = TextBox3.Text
itemform(showing, 2) = TextBox83.Text
itemform(showing, 3) = arrays(16, C)
If arrays(13, C) = "Garment" Then itemform(showing, 4) = Format(Round(TextBox83.Text + (TextBox83.Text * arrays(20, C)), 2), "fixed")
If arrays(13, C) = "Garment" Then itemform(showing, 5) = Format(Round(arrays(16, C) + (arrays(16, C) * arrays(20, C)), 2), "fixed")
itemform(showing, 6) = arrays(12, C)
itemform(showing, 7) = arrays(21, C)
itemform(showing, 8) = arrays(22, C)
itemform(showing, 9) = arrays(23, C)
itemform(showing, 10) = ComboBox4.Value
itemform(showing, 12) = ComboBox7.Value
itemform(showing, 13) = OptionButton4.Value
itemform(showing, 14) = ComboBox6.Value
itemform(showing, 15) = C
If lastitem < showing Then lastitem = showing
End If
If which = 2 And ComboBox10.ListIndex >= 0 Then
C = list2(ComboBox10.ListIndex)
TextBox85.Text = arrays(2, C)
TextBox4.Text = arrays(14, C)
If ComboBox16.ListIndex = 0 Then
If ComboBox8.ListIndex = 0 Then TextBox86.Text = arrays(3, C)
If ComboBox8.ListIndex = 1 Then TextBox86.Text = arrays(4, C)
If ComboBox8.ListIndex = 2 Then TextBox86.Text = arrays(5, C)
End If
If ComboBox16.ListIndex = 1 Then
If ComboBox8.ListIndex = 0 Then TextBox86.Text = arrays(6, C)
If ComboBox8.ListIndex = 1 Then TextBox86.Text = arrays(7, C)
If ComboBox8.ListIndex = 2 Then TextBox86.Text = arrays(8, C)
End If
If ComboBox16.ListIndex = 2 Then
If ComboBox8.ListIndex = 0 Then TextBox86.Text = arrays(9, C)
If ComboBox8.ListIndex = 1 Then TextBox86.Text = arrays(10, C)
If ComboBox8.ListIndex = 2 Then TextBox86.Text = arrays(11, C)
End If
If freqchange <> 1 Then
ComboBox9.Clear
colour = Split(arrays(18, C), ",")
If InStr(arrays(18, C), ",") = 0 Then
ComboBox9.AddItem arrays(18, C)
ComboBox9.ListIndex = 0
TextBox4.Text = Left(TextBox4.Text, 8) & "-" & Left(ComboBox9.Value, 3)
Else
For i = LBound(colour) To UBound(colour)
ComboBox9.AddItem colour(i)
Next i
End If
End If
itemform(showing + 1, 0) = arrays(1, C)
itemform(showing + 1, 1) = TextBox4.Text
itemform(showing + 1, 2) = TextBox86.Text
itemform(showing + 1, 3) = arrays(16, C)
If arrays(13, C) = "Garment" Then itemform(showing + 1, 4) = Format(Round(TextBox86.Text + (TextBox86.Text * arrays(20, C)), 2), "fixed")
If arrays(13, C) = "Garment" Then itemform(showing + 1, 5) = Format(Round(arrays(16, C) + (arrays(16, C) * arrays(20, C)), 2), "fixed")
itemform(showing + 1, 6) = arrays(12, C)
itemform(showing + 1, 7) = arrays(21, C)
itemform(showing + 1, 8) = arrays(22, C)
itemform(showing + 1, 9) = arrays(23, C)
itemform(showing + 1, 10) = ComboBox9.Value
itemform(showing + 1, 12) = ComboBox8.Value
itemform(showing + 1, 13) = OptionButton6.Value
itemform(showing + 1, 14) = ComboBox11.Value
itemform(showing + 1, 15) = C
If lastitem < showing + 1 Then lastitem = showing + 1
End If
If which = 3 And ComboBox14.ListIndex >= 0 Then
C = list3(ComboBox14.ListIndex)
TextBox90.Text = arrays(2, C)
TextBox6.Text = arrays(14, C)
If ComboBox16.ListIndex = 0 Then
If ComboBox12.ListIndex = 0 Then TextBox89.Text = arrays(3, C)
If ComboBox12.ListIndex = 1 Then TextBox89.Text = arrays(4, C)
If ComboBox12.ListIndex = 2 Then TextBox89.Text = arrays(5, C)
End If
If ComboBox16.ListIndex = 1 Then
If ComboBox12.ListIndex = 0 Then TextBox89.Text = arrays(6, C)
If ComboBox12.ListIndex = 1 Then TextBox89.Text = arrays(7, C)
If ComboBox12.ListIndex = 2 Then TextBox89.Text = arrays(8, C)
End If
If ComboBox16.ListIndex = 2 Then
If ComboBox12.ListIndex = 0 Then TextBox89.Text = arrays(9, C)
If ComboBox12.ListIndex = 1 Then TextBox89.Text = arrays(10, C)
If ComboBox12.ListIndex = 2 Then TextBox89.Text = arrays(11, C)
End If
If freqchange <> 1 Then
ComboBox13.Clear
colour = Split(arrays(18, C), ",")
If InStr(arrays(18, C), ",") = 0 Then
ComboBox13.AddItem arrays(18, C)
ComboBox13.ListIndex = 0
TextBox6.Text = Left(TextBox6.Text, 8) & "-" & Left(ComboBox13.Value, 3)
Else
For i = LBound(colour) To UBound(colour)
ComboBox13.AddItem colour(i)
Next i
End If
End If
itemform(showing + 2, 0) = arrays(1, C)
itemform(showing + 2, 1) = TextBox6.Text
itemform(showing + 2, 2) = TextBox89.Text
itemform(showing + 2, 3) = arrays(16, C)
If arrays(13, C) = "Garment" Then itemform(showing + 2, 4) = Format(Round(TextBox89.Text + (TextBox89.Text * arrays(20, C)), 2), "fixed")
If arrays(13, C) = "Garment" Then itemform(showing + 2, 5) = Format(Round(arrays(16, C) + (arrays(16, C) * arrays(20, C)), 2), "Fixed")
itemform(showing + 2, 6) = arrays(12, C)
itemform(showing + 2, 7) = arrays(21, C)
itemform(showing + 2, 8) = arrays(22, C)
itemform(showing + 2, 9) = arrays(23, C)
itemform(showing + 2, 10) = ComboBox13.Value
itemform(showing + 2, 12) = ComboBox12.Value
itemform(showing + 2, 13) = OptionButton8.Value
itemform(showing + 2, 14) = ComboBox15.Value
itemform(showing + 2, 15) = C
If lastitem < showing + 2 Then lastitem = showing + 2
End If
Call calculate
End If
End If
End Sub
Private Sub ComboBox6_Change()
If tempthing = 1 Then tempthing = 0: Exit Sub
If loaded = 1 Then
If scrolling <> 1 And itemform(showing, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
tempthing = 1
ComboBox6.Value = itemform(showing, 14)
Exit Sub
End If
Call filter(1)
End If
End Sub
Private Sub ComboBox9_Change()
If scrolling <> 1 Then
itemform(showing + 1, 10) = ComboBox9.Value
If ComboBox9.Value <> "" Then TextBox4.Text = Left(TextBox4.Text, 8) & "-" & Left(ComboBox9.Value, 3)
itemform(showing + 1, 1) = TextBox4.Text
End If
End Sub
Private Sub CommandButton1_Click()
clicked = showing
UserForm2.Show
End Sub
Private Sub CommandButton10_Click()
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
If a < 999 Then a = a + 1
Label96.Caption = a
TextBox70.Text = users(a, 0)
TextBox71.Text = users(a, 1)
TextBox72.Text = users(a, 2)
TextBox73.Text = users(a, 3)
TextBox74.Text = users(a, 4)
ComboBox20.Value = users(a, 5)
TextBox58.Text = users(a, 6)
TextBox61.Text = users(a, 7)
TextBox64.Text = users(a, 8)
TextBox65.Text = users(a, 9)
If users(a, 10) <> "" Then CheckBox1.Value = users(a, 10)
If users(a, 11) <> "" Then CheckBox4.Value = users(a, 11)
If users(a, 12) <> "" Then CheckBox7.Value = users(a, 12)
ComboBox21.Value = users(a, 13)
TextBox59.Text = users(a, 14)
TextBox62.Text = users(a, 15)
TextBox66.Text = users(a, 16)
TextBox67.Text = users(a, 17)
If users(a, 18) <> "" Then CheckBox2.Value = users(a, 18)
If users(a, 19) <> "" Then CheckBox5.Value = users(a, 19)
If users(a, 20) <> "" Then CheckBox8.Value = users(a, 20)
ComboBox22.Value = users(a, 21)
TextBox60.Text = users(a, 22)
TextBox63.Text = users(a, 23)
TextBox68.Text = users(a, 24)
TextBox69.Text = users(a, 25)
If users(a, 26) <> "" Then CheckBox3.Value = users(a, 26)
If users(a, 27) <> "" Then CheckBox6.Value = users(a, 27)
If users(a, 28) <> "" Then CheckBox9.Value = users(a, 28)
If lastuser < a Then lastuser = a
TextBox70.SetFocus
End Sub
Private Sub CommandButton11_Click()
dpc = ComboBox33.Value
taxrate = TextBox117.Text
salesrep = TextBox112.Text
If MsgBox("Are you sure?", vbYesNo) = 6 Then
Unload Me
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
Erase itemform, users, setrefill, usergarment, list1, list2, list3
lastuser = 1
lastitem = 3
lastrefill = 3
UserForm1.Show
End If
End Sub
Private Sub CommandButton12_Click()
TextBox34.Text = TextBox24.Text
TextBox35.Text = TextBox26.Text
TextBox36.Text = TextBox27.Text
TextBox37.Text = TextBox28.Text
TextBox38.Text = TextBox29.Text
TextBox115.Text = TextBox114.Text
End Sub
Private Sub CommandButton13_Click()
On Error GoTo errorhandler2
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
If lastrefill = "" Then lastrefill = 3
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
mystr = "A1:Q99"
Sheets("Data1").Range(mystr) = itemform
mystr = "A1:E99"
Sheets("Data2").Range(mystr) = setrefill
mystr = "A1:AF999"
Sheets("Data3").Range(mystr) = users
Sheets("Data4").Range("A1").Value = TextBox9.Text
Sheets("Data4").Range("A2").Value = TextBox11.Text
Sheets("Data4").Range("A3").Value = TextBox20.Text
Sheets("Data4").Range("A4").Value = TextBox13.Text
Sheets("Data4").Range("A5").Value = TextBox21.Text
Sheets("Data4").Range("A6").Value = ComboBox17.Value
Sheets("Data4").Range("A7").Value = ComboBox18.Value
Sheets("Data4").Range("A8").Value = TextBox8.Text
Sheets("Data4").Range("A9").Value = ComboBox16.Value
Sheets("Data4").Range("A10").Value = TextBox22.Text
Sheets("Data4").Range("A11").Value = TextBox23.Text
Sheets("Data4").Range("A12").Value = TextBox24.Text
Sheets("Data4").Range("A13").Value = TextBox25.Text
Sheets("Data4").Range("A14").Value = TextBox26.Text
Sheets("Data4").Range("A15").Value = TextBox27.Text
Sheets("Data4").Range("A16").Value = TextBox28.Text
Sheets("Data4").Range("A17").Value = TextBox29.Text
Sheets("Data4").Range("A18").Value = TextBox30.Text
Sheets("Data4").Range("A19").Value = TextBox31.Text
Sheets("Data4").Range("A20").Value = TextBox32.Text
Sheets("Data4").Range("A21").Value = TextBox33.Text
Sheets("Data4").Range("A22").Value = TextBox34.Text
Sheets("Data4").Range("A23").Value = TextBox35.Text
Sheets("Data4").Range("A24").Value = TextBox36.Text
Sheets("Data4").Range("A25").Value = TextBox37.Text
Sheets("Data4").Range("A26").Value = TextBox38.Text
Sheets("Data4").Range("A27").Value = TextBox39.Text
Sheets("Data4").Range("A28").Value = TextBox40.Text
Sheets("Data4").Range("A29").Value = TextBox41.Text
Sheets("Data4").Range("A30").Value = TextBox42.Text
Sheets("Data4").Range("A31").Value = TextBox43.Text
Sheets("Data4").Range("A32").Value = TextBox44.Text
Sheets("Data4").Range("A33").Value = TextBox45.Text
Sheets("Data4").Range("A34").Value = TextBox46.Text
Sheets("Data4").Range("A35").Value = TextBox47.Text
Sheets("Data4").Range("A36").Value = TextBox48.Text
Sheets("Data4").Range("A37").Value = TextBox49.Text
Sheets("Data4").Range("A38").Value = TextBox50.Text
Sheets("Data4").Range("A39").Value = TextBox51.Text
Sheets("Data4").Range("A40").Value = TextBox52.Text
Sheets("Data4").Range("A41").Value = TextBox53.Text
Sheets("Data4").Range("A42").Value = TextBox55.Text
Sheets("Data4").Range("A43").Value = TextBox56.Text
Sheets("Data4").Range("A44").Value = TextBox57.Text
Sheets("Data4").Range("A45").Value = ComboBox19.Value
Sheets("Data4").Range("A46").Value = ComboBox26.Value
Sheets("Data4").Range("A47").Value = TextBox81.Text
Sheets("Data4").Range("A48").Value = TextBox91.Text
Sheets("Data4").Range("A49").Value = lastuser
Sheets("Data4").Range("A50").Value = lastitem
Sheets("Data4").Range("A51").Value = lastrefill
Sheets("Data4").Range("A52").Value = TextBox114.Text
Sheets("Data4").Range("A53").Value = TextBox115.Text
Sheets("Data4").Range("A54").Value = TextBox92.Text
Sheets("Data4").Range("A55").Value = TextBox93.Text
Sheets("Data4").Range("A56").Value = TextBox94.Text
Sheets("Data4").Range("A57").Value = TextBox95.Text
Sheets("Data4").Range("A58").Value = TextBox96.Text
Sheets("Data4").Range("A59").Value = TextBox97.Text
Sheets("Data4").Range("A60").Value = TextBox98.Text
Sheets("Data4").Range("A61").Value = TextBox99.Text
Sheets("Data4").Range("A62").Value = TextBox100.Text
Sheets("Data4").Range("A63").Value = TextBox101.Text
Sheets("Data4").Range("A64").Value = TextBox102.Text
Sheets("Data4").Range("A65").Value = TextBox103.Text
Sheets("Data4").Range("A66").Value = TextBox104.Text
Sheets("Data4").Range("A67").Value = TextBox105.Text
Sheets("Data4").Range("A68").Value = TextBox106.Text
Sheets("Data4").Range("A69").Value = TextBox107.Text
Sheets("Data4").Range("A70").Value = TextBox108.Text
Sheets("Data4").Range("A71").Value = TextBox117.Text
Sheets("Data4").Range("A72").Value = TextBox112.Text
Sheets("Data4").Range("A73").Value = ComboBox33.Value
Sheets("Data4").Range("A74").Value = TextBox118.Text
Sheets("Data4").Range("A75").Value = TextBox119.Text
Dim oEmbFile As Object
Application.DisplayAlerts = False
Set oEmbFile = ThisWorkbook.Sheets("SA").OLEObjects(1)
    oEmbFile.Verb Verb:=xlPrimary
Set oEmbFile = Nothing
Application.DisplayAlerts = True
Dim wrddoc As Document
Application.Wait (Now + TimeValue("00:00:04"))
Set objWord = GetObject(, "Word.Application")
objWord.Visible = True
Set wrddoc = objWord.Documents("Document in " & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)))
wrddoc.Shapes("Text Box 23").TextFrame.TextRange = UserForm1.TextBox8.Text
wrddoc.Shapes("Text Box 29").TextFrame.TextRange = UserForm1.TextBox109.Text
wrddoc.Shapes("Text Box 32").TextFrame.TextRange = UserForm1.TextBox110.Text
wrddoc.Shapes("Text Box 35").TextFrame.TextRange = UserForm1.TextBox92.Text
wrddoc.Shapes("Text Box 40").TextFrame.TextRange = UserForm1.TextBox92.Text
wrddoc.Shapes("Text Box 43").TextFrame.TextRange = UserForm1.TextBox103.Text
wrddoc.Shapes("Text Box 44").TextFrame.TextRange = UserForm1.TextBox104.Text
wrddoc.Shapes("Text Box 45").TextFrame.TextRange = UserForm1.TextBox105.Text & "%"
wrddoc.Shapes("Text Box 46").TextFrame.TextRange = UserForm1.TextBox108.Text
wrddoc.Shapes("Text Box 47").TextFrame.TextRange = UserForm1.TextBox102.Text

wrddoc.Tables(3).Rows(2).Cells(2).Range.Text = UserForm1.TextBox93.Text
wrddoc.Tables(3).Rows(3).Cells(2).Range.Text = UserForm1.TextBox94.Text
wrddoc.Tables(3).Rows(4).Cells(2).Range.Text = UserForm1.TextBox95.Text
wrddoc.Tables(3).Rows(5).Cells(2).Range.Text = UserForm1.TextBox96.Text

wrddoc.Tables(3).Rows(2).Cells(4).Range.Text = UserForm1.TextBox97.Text & "%"
wrddoc.Tables(3).Rows(3).Cells(4).Range.Text = UserForm1.TextBox111.Text
wrddoc.Tables(3).Rows(4).Cells(4).Range.Text = UserForm1.TextBox106.Text
wrddoc.Tables(3).Rows(5).Cells(4).Range.Text = UserForm1.TextBox107.Text

wrddoc.Tables(3).Rows(2).Cells(6).Range.Text = UserForm1.TextBox98.Text
wrddoc.Tables(3).Rows(3).Cells(6).Range.Text = UserForm1.TextBox99.Text & "%"
wrddoc.Tables(3).Rows(4).Cells(6).Range.Text = UserForm1.TextBox100.Text & "%"
wrddoc.Tables(3).Rows(5).Cells(6).Range.Text = UserForm1.TextBox101.Text & "%"

wrddoc.Tables(1).Rows(1).Cells(2).Range.Text = "G&K SERVICES CANADA INC." & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 3, False) & " " & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 4, False) & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 5, False)
a = 1
b = 3
Do While a < lastitem + 1 And b < 13
If itemform(a, 0) <> "" And itemform(a, 11) > 0 And itemform(a, 13) = True Then
wrddoc.Tables(2).Rows(b).Cells(1).Range.Text = itemform(a, 0) & " " & itemform(a, 1)
wrddoc.Tables(2).Rows(b).Cells(5).Range.Text = itemform(a, 12)
wrddoc.Tables(2).Rows(b).Cells(4).Range.Text = itemform(a, 2)
wrddoc.Tables(2).Rows(b).Cells(6).Range.Text = itemform(a, 3)
If arrays(15, itemform(a, 15)) = "US" Then wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = itemform(a, 11) / 2: wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = itemform(a, 11): GoTo here
If Left(itemform(a, 1), 4) > 1999 And Left(itemform(a, 1), 4) < 2100 Then wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = itemform(a, 11): wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = itemform(a, 11) * 2: GoTo here
If arrays(15, itemform(a, 15)) <> "CI" Then wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = itemform(a, 11): wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = itemform(a, 11): GoTo here
wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = TextBox118.Text: wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = TextBox119.Text
here:
b = b + 1
End If
a = a + 1
Loop
Do While b < 13
wrddoc.Tables(2).Rows(b).Cells(1).Range.Text = ""
wrddoc.Tables(2).Rows(b).Cells(5).Range.Text = ""
wrddoc.Tables(2).Rows(b).Cells(4).Range.Text = ""
wrddoc.Tables(2).Rows(b).Cells(6).Range.Text = ""
wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = ""
wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = ""
b = b + 1
Loop
Set wrddoc = Nothing
Exit Sub
errorhandler2:
Set wrddoc = objWord.Documents("Document in " & ThisWorkbook.Name)
Resume Next
End Sub
Private Sub CommandButton14_Click()
On Error GoTo errorhandler
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
If lastrefill = "" Then lastrefill = 3
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
mystr = "A1:Q99"
Sheets("Data1").Range(mystr) = itemform
mystr = "A1:E99"
Sheets("Data2").Range(mystr) = setrefill
mystr = "A1:AF999"
Sheets("Data3").Range(mystr) = users
Sheets("Data4").Range("A1").Value = TextBox9.Text
Sheets("Data4").Range("A2").Value = TextBox11.Text
Sheets("Data4").Range("A3").Value = TextBox20.Text
Sheets("Data4").Range("A4").Value = TextBox13.Text
Sheets("Data4").Range("A5").Value = TextBox21.Text
Sheets("Data4").Range("A6").Value = ComboBox17.Value
Sheets("Data4").Range("A7").Value = ComboBox18.Value
Sheets("Data4").Range("A8").Value = TextBox8.Text
Sheets("Data4").Range("A9").Value = ComboBox16.Value
Sheets("Data4").Range("A10").Value = TextBox22.Text
Sheets("Data4").Range("A11").Value = TextBox23.Text
Sheets("Data4").Range("A12").Value = TextBox24.Text
Sheets("Data4").Range("A13").Value = TextBox25.Text
Sheets("Data4").Range("A14").Value = TextBox26.Text
Sheets("Data4").Range("A15").Value = TextBox27.Text
Sheets("Data4").Range("A16").Value = TextBox28.Text
Sheets("Data4").Range("A17").Value = TextBox29.Text
Sheets("Data4").Range("A18").Value = TextBox30.Text
Sheets("Data4").Range("A19").Value = TextBox31.Text
Sheets("Data4").Range("A20").Value = TextBox32.Text
Sheets("Data4").Range("A21").Value = TextBox33.Text
Sheets("Data4").Range("A22").Value = TextBox34.Text
Sheets("Data4").Range("A23").Value = TextBox35.Text
Sheets("Data4").Range("A24").Value = TextBox36.Text
Sheets("Data4").Range("A25").Value = TextBox37.Text
Sheets("Data4").Range("A26").Value = TextBox38.Text
Sheets("Data4").Range("A27").Value = TextBox39.Text
Sheets("Data4").Range("A28").Value = TextBox40.Text
Sheets("Data4").Range("A29").Value = TextBox41.Text
Sheets("Data4").Range("A30").Value = TextBox42.Text
Sheets("Data4").Range("A31").Value = TextBox43.Text
Sheets("Data4").Range("A32").Value = TextBox44.Text
Sheets("Data4").Range("A33").Value = TextBox45.Text
Sheets("Data4").Range("A34").Value = TextBox46.Text
Sheets("Data4").Range("A35").Value = TextBox47.Text
Sheets("Data4").Range("A36").Value = TextBox48.Text
Sheets("Data4").Range("A37").Value = TextBox49.Text
Sheets("Data4").Range("A38").Value = TextBox50.Text
Sheets("Data4").Range("A39").Value = TextBox51.Text
Sheets("Data4").Range("A40").Value = TextBox52.Text
Sheets("Data4").Range("A41").Value = TextBox53.Text
Sheets("Data4").Range("A42").Value = TextBox55.Text
Sheets("Data4").Range("A43").Value = TextBox56.Text
Sheets("Data4").Range("A44").Value = TextBox57.Text
Sheets("Data4").Range("A45").Value = ComboBox19.Value
Sheets("Data4").Range("A46").Value = ComboBox26.Value
Sheets("Data4").Range("A47").Value = TextBox81.Text
Sheets("Data4").Range("A48").Value = TextBox91.Text
Sheets("Data4").Range("A49").Value = lastuser
Sheets("Data4").Range("A50").Value = lastitem
Sheets("Data4").Range("A51").Value = lastrefill
Sheets("Data4").Range("A52").Value = TextBox114.Text
Sheets("Data4").Range("A53").Value = TextBox115.Text
Sheets("Data4").Range("A54").Value = TextBox92.Text
Sheets("Data4").Range("A55").Value = TextBox93.Text
Sheets("Data4").Range("A56").Value = TextBox94.Text
Sheets("Data4").Range("A57").Value = TextBox95.Text
Sheets("Data4").Range("A58").Value = TextBox96.Text
Sheets("Data4").Range("A59").Value = TextBox97.Text
Sheets("Data4").Range("A60").Value = TextBox98.Text
Sheets("Data4").Range("A61").Value = TextBox99.Text
Sheets("Data4").Range("A62").Value = TextBox100.Text
Sheets("Data4").Range("A63").Value = TextBox101.Text
Sheets("Data4").Range("A64").Value = TextBox102.Text
Sheets("Data4").Range("A65").Value = TextBox103.Text
Sheets("Data4").Range("A66").Value = TextBox104.Text
Sheets("Data4").Range("A67").Value = TextBox105.Text
Sheets("Data4").Range("A68").Value = TextBox106.Text
Sheets("Data4").Range("A69").Value = TextBox107.Text
Sheets("Data4").Range("A70").Value = TextBox108.Text

Sheets("Data4").Range("A71").Value = TextBox117.Text
Sheets("Data4").Range("A72").Value = TextBox112.Text
Sheets("Data4").Range("A73").Value = ComboBox33.Value
Sheets("Data4").Range("A74").Value = TextBox118.Text
Sheets("Data4").Range("A75").Value = TextBox119.Text
Dim oEmbFile As Object
Application.DisplayAlerts = False
Set oEmbFile = ThisWorkbook.Sheets("SA").OLEObjects(2)
    oEmbFile.Verb Verb:=xlPrimary
Set oEmbFile = Nothing
Application.DisplayAlerts = True
Dim wrddoc As Document
Application.Wait (Now + TimeValue("00:00:04"))
   Set objWord = GetObject(, "Word.Application")
   objWord.Visible = True
Set wrddoc = objWord.Documents("Document in " & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)))
wrddoc.Shapes("Text Box 5").TextFrame.TextRange = UserForm1.TextBox8.Text
wrddoc.Shapes("Text Box 6").TextFrame.TextRange = UserForm1.TextBox109.Text
wrddoc.Shapes("Text Box 8").TextFrame.TextRange = UserForm1.TextBox92.Text
wrddoc.Shapes("Text Box 9").TextFrame.TextRange = UserForm1.TextBox102.Text
wrddoc.Shapes("Text Box 10").TextFrame.TextRange = UserForm1.TextBox110.Text
wrddoc.Tables(1).Rows(1).Cells(2).Range.Text = "G&K SERVICES CANADA INC." & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 3, False) & " " & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 4, False) & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 5, False)
wrddoc.Tables(4).Rows(2).Cells(2).Range.Text = UserForm1.TextBox93.Text
wrddoc.Tables(4).Rows(3).Cells(2).Range.Text = UserForm1.TextBox94.Text
wrddoc.Tables(4).Rows(4).Cells(2).Range.Text = UserForm1.TextBox95.Text
wrddoc.Tables(4).Rows(5).Cells(2).Range.Text = UserForm1.TextBox96.Text
wrddoc.Tables(4).Rows(2).Cells(4).Range.Text = UserForm1.TextBox97.Text & "%"
wrddoc.Tables(4).Rows(3).Cells(4).Range.Text = UserForm1.TextBox111.Text
wrddoc.Tables(4).Rows(4).Cells(4).Range.Text = UserForm1.TextBox106.Text
wrddoc.Tables(4).Rows(5).Cells(4).Range.Text = UserForm1.TextBox107.Text
wrddoc.Tables(4).Rows(2).Cells(6).Range.Text = UserForm1.TextBox98.Text
wrddoc.Tables(4).Rows(3).Cells(6).Range.Text = UserForm1.TextBox99.Text & "%"
wrddoc.Tables(4).Rows(4).Cells(6).Range.Text = UserForm1.TextBox100.Text & "%"
wrddoc.Tables(4).Rows(5).Cells(6).Range.Text = UserForm1.TextBox101.Text & "%"

a = 1
b = 3
Do While a < lastitem + 1 And b < 32
If itemform(a, 0) <> "" And itemform(a, 11) > 0 And itemform(a, 13) = True Then
If b > 12 Then
wrddoc.Tables(2).Rows(b - 10).Cells(1).Range.Text = itemform(a, 0) & " " & itemform(a, 1)
wrddoc.Tables(2).Rows(b - 10).Cells(5).Range.Text = itemform(a, 12)
wrddoc.Tables(2).Rows(b - 10).Cells(4).Range.Text = itemform(a, 2)
wrddoc.Tables(2).Rows(b - 10).Cells(6).Range.Text = itemform(a, 3)
If arrays(15, itemform(a, 15)) = "US" Then wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = itemform(a, 11) / 2: wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = itemform(a, 11): GoTo here
If Left(itemform(a, 1), 4) > 1999 And Left(itemform(a, 1), 4) < 2100 Then wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = itemform(a, 11): wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = itemform(a, 11) * 2: GoTo here
If arrays(15, itemform(a, 15)) <> "CI" Then wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = itemform(a, 11): wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = itemform(a, 11): GoTo here
wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = TextBox118.Text: wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = TextBox119.Text
End If
here:
b = b + 1
End If
a = a + 1
Loop
Do While b < 29
If b < 14 Then b = 13
wrddoc.Tables(2).Rows(b - 10).Cells(1).Range.Text = ""
wrddoc.Tables(2).Rows(b - 10).Cells(5).Range.Text = ""
wrddoc.Tables(2).Rows(b - 10).Cells(4).Range.Text = ""
wrddoc.Tables(2).Rows(b - 10).Cells(6).Range.Text = ""
wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = ""
wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = ""
b = b + 1
Loop
Set wrddoc = Nothing
Exit Sub
errorhandler:
Set wrddoc = objWord.Documents("Document in " & ThisWorkbook.Name)
Resume Next
End Sub

Private Sub CommandButton15_Click()
MsgBox Word.Application.Documents("Document in " & ThisWorkbook.Name)
End Sub

Private Sub CommandButton2_Click()
clicked = showing + 1
UserForm2.Show
End Sub
Private Sub CommandButton3_Click()
clicked = showing + 2
UserForm2.Show
End Sub
Public Function valueofitem(which As Integer)
If which = 1 Then If itemform(showing, 14) = "" Then valueofitem = "ALL" Else valueofitem = itemform(showing, 14)
If which = 2 Then If itemform(showing + 1, 14) = "" Then valueofitem = "ALL" Else valueofitem = itemform(showing + 1, 14)
If which = 3 Then If itemform(showing + 2, 14) = "" Then valueofitem = "ALL" Else valueofitem = itemform(showing + 2, 14)
End Function
Public Function valueofitem2(which As Integer)
If which = 1 Then If itemform(showing, 12) = "" Then valueofitem2 = "WEEKLY" Else valueofitem2 = itemform(showing, 12)
If which = 2 Then If itemform(showing + 1, 12) = "" Then valueofitem2 = "WEEKLY" Else valueofitem2 = itemform(showing + 1, 12)
If which = 3 Then If itemform(showing + 2, 12) = "" Then valueofitem2 = "WEEKLY" Else valueofitem2 = itemform(showing + 2, 12)
End Function
Private Sub CommandButton7_Click()
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
If lastrefill = "" Then lastrefill = 3
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
mystr = "A1:Q99"
Sheets("Data1").Range(mystr) = itemform
mystr = "A1:E99"
Sheets("Data2").Range(mystr) = setrefill
mystr = "A1:AF999"
Sheets("Data3").Range(mystr) = users
Sheets("Data4").Range("A1").Value = TextBox9.Text
Sheets("Data4").Range("A2").Value = TextBox11.Text
Sheets("Data4").Range("A3").Value = TextBox20.Text
Sheets("Data4").Range("A4").Value = TextBox13.Text
Sheets("Data4").Range("A5").Value = TextBox21.Text
Sheets("Data4").Range("A6").Value = ComboBox17.Value
Sheets("Data4").Range("A7").Value = ComboBox18.Value
Sheets("Data4").Range("A8").Value = TextBox8.Text
Sheets("Data4").Range("A9").Value = ComboBox16.Value
Sheets("Data4").Range("A10").Value = TextBox22.Text
Sheets("Data4").Range("A11").Value = TextBox23.Text
Sheets("Data4").Range("A12").Value = TextBox24.Text
Sheets("Data4").Range("A13").Value = TextBox25.Text
Sheets("Data4").Range("A14").Value = TextBox26.Text
Sheets("Data4").Range("A15").Value = TextBox27.Text
Sheets("Data4").Range("A16").Value = TextBox28.Text
Sheets("Data4").Range("A17").Value = TextBox29.Text
Sheets("Data4").Range("A18").Value = TextBox30.Text
Sheets("Data4").Range("A19").Value = TextBox31.Text
Sheets("Data4").Range("A20").Value = TextBox32.Text
Sheets("Data4").Range("A21").Value = TextBox33.Text
Sheets("Data4").Range("A22").Value = TextBox34.Text
Sheets("Data4").Range("A23").Value = TextBox35.Text
Sheets("Data4").Range("A24").Value = TextBox36.Text
Sheets("Data4").Range("A25").Value = TextBox37.Text
Sheets("Data4").Range("A26").Value = TextBox38.Text
Sheets("Data4").Range("A27").Value = TextBox39.Text
Sheets("Data4").Range("A28").Value = TextBox40.Text
Sheets("Data4").Range("A29").Value = TextBox41.Text
Sheets("Data4").Range("A30").Value = TextBox42.Text
Sheets("Data4").Range("A31").Value = TextBox43.Text
Sheets("Data4").Range("A32").Value = TextBox44.Text
Sheets("Data4").Range("A33").Value = TextBox45.Text
Sheets("Data4").Range("A34").Value = TextBox46.Text
Sheets("Data4").Range("A35").Value = TextBox47.Text
Sheets("Data4").Range("A36").Value = TextBox48.Text
Sheets("Data4").Range("A37").Value = TextBox49.Text
Sheets("Data4").Range("A38").Value = TextBox50.Text
Sheets("Data4").Range("A39").Value = TextBox51.Text
Sheets("Data4").Range("A40").Value = TextBox52.Text
Sheets("Data4").Range("A41").Value = TextBox53.Text
Sheets("Data4").Range("A42").Value = TextBox55.Text
Sheets("Data4").Range("A43").Value = TextBox56.Text
Sheets("Data4").Range("A44").Value = TextBox57.Text
Sheets("Data4").Range("A45").Value = ComboBox19.Value
Sheets("Data4").Range("A46").Value = ComboBox26.Value
Sheets("Data4").Range("A47").Value = TextBox81.Text
Sheets("Data4").Range("A48").Value = TextBox91.Text
Sheets("Data4").Range("A49").Value = lastuser
Sheets("Data4").Range("A50").Value = lastitem
Sheets("Data4").Range("A51").Value = lastrefill
Sheets("Data4").Range("A52").Value = TextBox114.Text
Sheets("Data4").Range("A53").Value = TextBox115.Text
Sheets("Data4").Range("A54").Value = TextBox92.Text
Sheets("Data4").Range("A55").Value = TextBox93.Text
Sheets("Data4").Range("A56").Value = TextBox94.Text
Sheets("Data4").Range("A57").Value = TextBox95.Text
Sheets("Data4").Range("A58").Value = TextBox96.Text
Sheets("Data4").Range("A59").Value = TextBox97.Text
Sheets("Data4").Range("A60").Value = TextBox98.Text
Sheets("Data4").Range("A61").Value = TextBox99.Text
Sheets("Data4").Range("A62").Value = TextBox100.Text
Sheets("Data4").Range("A63").Value = TextBox101.Text
Sheets("Data4").Range("A64").Value = TextBox102.Text
Sheets("Data4").Range("A65").Value = TextBox103.Text
Sheets("Data4").Range("A66").Value = TextBox104.Text
Sheets("Data4").Range("A67").Value = TextBox105.Text
Sheets("Data4").Range("A68").Value = TextBox106.Text
Sheets("Data4").Range("A69").Value = TextBox107.Text
Sheets("Data4").Range("A70").Value = TextBox108.Text
Sheets("Data4").Range("A71").Value = TextBox117.Text
Sheets("Data4").Range("A72").Value = TextBox112.Text
Sheets("Data4").Range("A73").Value = ComboBox33.Value
Sheets("Data4").Range("A74").Value = TextBox118.Text
Sheets("Data4").Range("A75").Value = TextBox119.Text
Sheets("Quote").EnableSelection = xlNoRestrictions
Sheets("Quote").Visible = True
Sheets("Quote").Copy
Sheets("Quote").Unprotect
ThisWorkbook.Sheets("Quote").Visible = False
UserForm1.Hide
Sheets("Quote").Cells(3, 6).Value = [TODAY()]
Sheets("Quote").Cells(9, 1).Value = TextBox8.Text
Sheets("Quote").Cells(3, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 3, False)
Sheets("Quote").Cells(4, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 4, False)
Sheets("Quote").Cells(5, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 5, False)
Sheets("Quote").Cells(7, 1).Value = "Prepared by: " & TextBox112.Text
Sheets("Quote").Cells(20, 6).Value = "=F19*" & TextBox117.Text & "%"
a = 1
newline = 14
Do While a < lastitem + 1
Sheets("Quote").Rows(newline).EntireRow.Insert
If itemform(a, 0) <> "" And itemform(a, 11) > 0 And itemform(a, 13) = True Then
Sheets("Quote").Cells(newline, 1).Value = itemform(a, 0)
Sheets("Quote").Cells(newline, 2).Value = itemform(a, 1)
Sheets("Quote").Cells(newline, 3).Value = itemform(a, 12)
Sheets("Quote").Cells(newline, 4).Value = itemform(a, 2)
Sheets("Quote").Cells(newline, 5).Value = itemform(a, 11)
Sheets("Quote").Cells(newline, 6).Value = "=D" & newline & "*E" & newline
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
If itemform(a, 6) > 0 Then
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Cells(newline, 1).Value = itemform(a, 0) & " AUTO-REPLACE " & itemform(a, 6) * 100 & "%"
Sheets("Quote").Cells(newline, 2).Value = itemform(a, 1)
Sheets("Quote").Cells(newline, 3).Value = itemform(a, 12)
Sheets("Quote").Cells(newline, 4).Value = itemform(a, 3)
Sheets("Quote").Cells(newline, 5).Value = Round(itemform(a, 11) * itemform(a, 6), 0)
Sheets("Quote").Cells(newline, 6).Value = "=D" & newline & "*E" & newline
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
End If
a = a + 1
Loop
If TextBox10.Text > 0 Then
Sheets("Quote").Cells(newline, 1).Value = "Invoice Minimum"
Sheets("Quote").Cells(newline, 6).Value = TextBox10.Text
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
If TextBox12.Text > 0 Then
Sheets("Quote").Cells(newline, 1).Value = "Energy Surcharge"
Sheets("Quote").Cells(newline, 6).Value = TextBox12.Text
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
If TextBox14.Text > 0 Then
Sheets("Quote").Cells(newline, 1).Value = "Environmental Surcharge"
Sheets("Quote").Cells(newline, 6).Value = TextBox14.Text
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
If TextBox16.Text > 0 Then
Sheets("Quote").Cells(newline, 1).Value = "Prep Guard"
Sheets("Quote").Cells(newline, 6).Value = TextBox16.Text
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
If TextBox15.Text > 0 Then
Sheets("Quote").Cells(newline, 1).Value = "Image Guard"
Sheets("Quote").Cells(newline, 6).Value = TextBox15.Text
newline = newline + 2
Sheets("Quote").Rows(newline).EntireRow.Insert
Sheets("Quote").Rows(newline).EntireRow.Insert
End If
End Sub
Private Sub CommandButton8_Click()
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
If lastrefill = "" Then lastrefill = 3
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
mystr = "A1:Q99"
Sheets("Data1").Range(mystr) = itemform
mystr = "A1:E99"
Sheets("Data2").Range(mystr) = setrefill
mystr = "A1:AF999"
Sheets("Data3").Range(mystr) = users
Sheets("Data4").Range("A1").Value = TextBox9.Text
Sheets("Data4").Range("A2").Value = TextBox11.Text
Sheets("Data4").Range("A3").Value = TextBox20.Text
Sheets("Data4").Range("A4").Value = TextBox13.Text
Sheets("Data4").Range("A5").Value = TextBox21.Text
Sheets("Data4").Range("A6").Value = ComboBox17.Value
Sheets("Data4").Range("A7").Value = ComboBox18.Value
Sheets("Data4").Range("A8").Value = TextBox8.Text
Sheets("Data4").Range("A9").Value = ComboBox16.Value
Sheets("Data4").Range("A10").Value = TextBox22.Text
Sheets("Data4").Range("A11").Value = TextBox23.Text
Sheets("Data4").Range("A12").Value = TextBox24.Text
Sheets("Data4").Range("A13").Value = TextBox25.Text
Sheets("Data4").Range("A14").Value = TextBox26.Text
Sheets("Data4").Range("A15").Value = TextBox27.Text
Sheets("Data4").Range("A16").Value = TextBox28.Text
Sheets("Data4").Range("A17").Value = TextBox29.Text
Sheets("Data4").Range("A18").Value = TextBox30.Text
Sheets("Data4").Range("A19").Value = TextBox31.Text
Sheets("Data4").Range("A20").Value = TextBox32.Text
Sheets("Data4").Range("A21").Value = TextBox33.Text
Sheets("Data4").Range("A22").Value = TextBox34.Text
Sheets("Data4").Range("A23").Value = TextBox35.Text
Sheets("Data4").Range("A24").Value = TextBox36.Text
Sheets("Data4").Range("A25").Value = TextBox37.Text
Sheets("Data4").Range("A26").Value = TextBox38.Text
Sheets("Data4").Range("A27").Value = TextBox39.Text
Sheets("Data4").Range("A28").Value = TextBox40.Text
Sheets("Data4").Range("A29").Value = TextBox41.Text
Sheets("Data4").Range("A30").Value = TextBox42.Text
Sheets("Data4").Range("A31").Value = TextBox43.Text
Sheets("Data4").Range("A32").Value = TextBox44.Text
Sheets("Data4").Range("A33").Value = TextBox45.Text
Sheets("Data4").Range("A34").Value = TextBox46.Text
Sheets("Data4").Range("A35").Value = TextBox47.Text
Sheets("Data4").Range("A36").Value = TextBox48.Text
Sheets("Data4").Range("A37").Value = TextBox49.Text
Sheets("Data4").Range("A38").Value = TextBox50.Text
Sheets("Data4").Range("A39").Value = TextBox51.Text
Sheets("Data4").Range("A40").Value = TextBox52.Text
Sheets("Data4").Range("A41").Value = TextBox53.Text
Sheets("Data4").Range("A42").Value = TextBox55.Text
Sheets("Data4").Range("A43").Value = TextBox56.Text
Sheets("Data4").Range("A44").Value = TextBox57.Text
Sheets("Data4").Range("A45").Value = ComboBox19.Value
Sheets("Data4").Range("A46").Value = ComboBox26.Value
Sheets("Data4").Range("A47").Value = TextBox81.Text
Sheets("Data4").Range("A48").Value = TextBox91.Text
Sheets("Data4").Range("A49").Value = lastuser
Sheets("Data4").Range("A50").Value = lastitem
Sheets("Data4").Range("A51").Value = lastrefill
Sheets("Data4").Range("A52").Value = TextBox114.Text
Sheets("Data4").Range("A53").Value = TextBox115.Text
Sheets("Data4").Range("A54").Value = TextBox92.Text
Sheets("Data4").Range("A55").Value = TextBox93.Text
Sheets("Data4").Range("A56").Value = TextBox94.Text
Sheets("Data4").Range("A57").Value = TextBox95.Text
Sheets("Data4").Range("A58").Value = TextBox96.Text
Sheets("Data4").Range("A59").Value = TextBox97.Text
Sheets("Data4").Range("A60").Value = TextBox98.Text
Sheets("Data4").Range("A61").Value = TextBox99.Text
Sheets("Data4").Range("A62").Value = TextBox100.Text
Sheets("Data4").Range("A63").Value = TextBox101.Text
Sheets("Data4").Range("A64").Value = TextBox102.Text
Sheets("Data4").Range("A65").Value = TextBox103.Text
Sheets("Data4").Range("A66").Value = TextBox104.Text
Sheets("Data4").Range("A67").Value = TextBox105.Text
Sheets("Data4").Range("A68").Value = TextBox106.Text
Sheets("Data4").Range("A69").Value = TextBox107.Text
Sheets("Data4").Range("A70").Value = TextBox108.Text
Sheets("Data4").Range("A71").Value = TextBox117.Text
Sheets("Data4").Range("A72").Value = TextBox112.Text
Sheets("Data4").Range("A73").Value = ComboBox33.Value
Sheets("Data4").Range("A74").Value = TextBox118.Text
Sheets("Data4").Range("A75").Value = TextBox119.Text
Sheets("Folder1").EnableSelection = xlNoRestrictions
Sheets("Folder1").Visible = True
Sheets("Folder1").Copy
Sheets("Folder1").Unprotect
ThisWorkbook.Sheets("Folder1").Visible = False
UserForm1.Hide
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Folder1")
If ComboBox16.ListIndex = 0 Then ws.Range("T5").Value = "1"
If ComboBox16.ListIndex = 1 Then ws.Range("T5").Value = "3"
If ComboBox16.ListIndex = 2 Then ws.Range("T5").Value = "5"
ws.Range("AI1").Value = Date
ws.Range("N1").Value = TextBox8.Text
ws.Range("A4").Value = TextBox22.Text
ws.Range("L4").Value = TextBox23.Text
ws.Range("A10").Value = TextBox24.Text
ws.Range("A13").Value = TextBox25.Text
ws.Range("A16").Value = TextBox26.Text
ws.Range("A19").Value = TextBox27.Text & " " & TextBox114.Text & " " & TextBox28.Text
ws.Range("V13").Value = TextBox114.Text
ws.Range("V17").Value = TextBox29.Text
ws.Range("V20").Value = TextBox30.Text
ws.Range("L10").Value = TextBox34.Text
ws.Range("L13").Value = TextBox35.Text
ws.Range("L16").Value = TextBox36.Text & " " & TextBox115.Text & " " & TextBox37.Text
ws.Range("L19").Value = TextBox38.Text
ws.Range("A24").Value = TextBox39.Text
ws.Range("I24").Value = TextBox40.Text
ws.Range("N25").Value = TextBox41.Text
ws.Range("L47").Value = TextBox42.Text
If TextBox42.Text = "" Then ws.Range("L47").Value = "N/A"
ws.Range("A29").Value = TextBox46.Text
If TextBox46.Text = "" Then ws.Range("A29").Value = "N/A"
ws.Range("O29").Value = TextBox47.Text
If TextBox47.Text = "" Then ws.Range("O29").Value = "N/A"
ws.Range("K40").Value = TextBox48.Text
ws.Range("N40").Value = TextBox49.Text
ws.Range("Q40").Value = TextBox50.Text
ws.Range("U40").Value = TextBox51.Text
If TextBox51.Text = "" Then ws.Range("U40").Value = "NO DUNS"
ws.Range("F43").Value = TextBox52.Text
ws.Range("L43").Value = TextBox53.Text
ws.Range("W48").Value = ComboBox19.Value
ws.Range("AD48").Value = ComboBox26.Value
ws.Range("AI3").Value = TextBox81.Text
ws.Range("A35").Value = TextBox31.Text
ws.Range("A36").Value = TextBox32.Text
ws.Range("A37").Value = TextBox33.Text
ws.Range("T35").Value = TextBox43.Text
ws.Range("T36").Value = TextBox44.Text
ws.Range("T37").Value = TextBox45.Text
If IsNumeric(TextBox55.Text) = True Then ws.Range("AA35").Value = TextBox55.Text / 100
If IsNumeric(TextBox56.Text) = True Then ws.Range("AA36").Value = TextBox56.Text / 100
If IsNumeric(TextBox57.Text) = True Then ws.Range("AA37").Value = TextBox57.Text / 100
ws.Range("A43").Value = TextBox9.Text
ws.Range("V8").Value = TextBox11.Text
ws.Range("AD8").Value = TextBox20.Text & "%"
ws.Range("V4").Value = TextBox13.Text
ws.Range("AD4").Value = TextBox21.Text & "%"
ws.Range("A51").Value = TextBox91.Text
ThisWorkbook.Sheets("Folder2").EnableSelection = xlNoRestrictions
ThisWorkbook.Sheets("Folder2").Visible = True
ThisWorkbook.Sheets("Folder2").Copy
ThisWorkbook.Sheets("Folder2").Unprotect
ThisWorkbook.Sheets("Folder2").Visible = False
newcolumn = 6
newrow = 0
itemref = 0
wsno = 1
anotherfour = 0
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Folder2")
bbb = 1
userrw = 30
usercol = 3
ccc = 1
ddd = 0
Dim iitem(999) As Integer
Do While ccc < lastitem + 1
If itemform(ccc, 13) = True And itemform(ccc, 0) <> "" Then
ddd = ddd + 1
iitem(ccc) = ddd
End If
ccc = ccc + 1
Loop
Do While bbb < 13
ws.Cells(userrw, usercol) = users(bbb, 0)
ws.Cells(userrw + 1, usercol) = users(bbb, 1)
ws.Cells(userrw, usercol + 6) = users(bbb, 2)
ws.Cells(userrw, usercol + 8) = users(bbb, 3)
ws.Cells(userrw, usercol + 10) = users(bbb, 4)
If users(bbb, 5) <> "" Then ws.Cells(userrw, usercol + 15) = iitem(Left(users(bbb, 5), 2))
ws.Cells(userrw, usercol + 17) = users(bbb, 8)
ws.Cells(userrw, usercol + 18) = users(bbb, 9)
If ws.Cells(userrw, usercol + 15) <> "" Then ws.Cells(userrw, usercol + 19) = "A"
ws.Cells(userrw, usercol + 21) = users(bbb, 7)
ws.Cells(userrw, usercol + 27) = users(bbb, 6)
If users(bbb, 10) = True Then ws.Cells(userrw, usercol + 24) = "Y"
If users(bbb, 11) = True Then ws.Cells(userrw, usercol + 25) = "Y"
If users(bbb, 12) = True Then ws.Cells(userrw, usercol + 26) = "Y"
If users(bbb, 13) <> "" Then ws.Cells(userrw + 1, usercol + 15) = iitem(Left(users(bbb, 13), 2))
ws.Cells(userrw + 1, usercol + 17) = users(bbb, 16)
ws.Cells(userrw + 1, usercol + 18) = users(bbb, 17)
If ws.Cells(userrw + 1, usercol + 15) <> "" Then ws.Cells(userrw + 1, usercol + 19) = "A"
ws.Cells(userrw + 1, usercol + 21) = users(bbb, 15)
ws.Cells(userrw + 1, usercol + 27) = users(bbb, 14)
If users(bbb, 18) = True Then ws.Cells(userrw + 1, usercol + 24) = "Y"
If users(bbb, 19) = True Then ws.Cells(userrw + 1, usercol + 25) = "Y"
If users(bbb, 20) = True Then ws.Cells(userrw + 1, usercol + 26) = "Y"
If users(bbb, 21) <> "" Then ws.Cells(userrw + 2, usercol + 15) = iitem(Left(users(bbb, 21), 2))
ws.Cells(userrw + 2, usercol + 17) = users(bbb, 24)
ws.Cells(userrw + 2, usercol + 18) = users(bbb, 25)
If ws.Cells(userrw + 2, usercol + 15) <> "" Then ws.Cells(userrw + 2, usercol + 19) = "A"
ws.Cells(userrw + 2, usercol + 21) = users(bbb, 23)
ws.Cells(userrw + 2, usercol + 27) = users(bbb, 22)
If users(bbb, 26) = True Then ws.Cells(userrw + 2, usercol + 24) = "Y"
If users(bbb, 27) = True Then ws.Cells(userrw + 2, usercol + 25) = "Y"
If users(bbb, 28) = True Then ws.Cells(userrw + 2, usercol + 26) = "Y"
bbb = bbb + 1
userrw = userrw + 3
If userrw = 48 Then usercol = 37: userrw = 30
Loop
a = 1
Do While a < lastitem + 1
If itemform(a, 13) = True And itemform(a, 0) <> "" Then
itemref = itemref + 1
If wsno = 2 And newcolumn = 6 And newrow = 0 And anotherfour = 0 Then
ThisWorkbook.Sheets("Folder3").EnableSelection = xlNoRestrictions
ThisWorkbook.Sheets("Folder3").Visible = True
ThisWorkbook.Sheets("Folder3").Copy
ThisWorkbook.Sheets("Folder3").Unprotect
ThisWorkbook.Sheets("Folder3").Visible = False
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Folder3")
End If
If anotherfour = 1 Then
Dim rgsource As Range
Dim rgdest
Set rgsource = ws.Range("A1:A25")
Set rgdest = ws.Range("A" & newrow + 27)
rgsource.EntireRow.Copy (rgdest)
ws.Range("F" & newrow + 28, "AC" & newrow + 51).ClearContents
ws.Cells(newrow + 27, 6) = itemref
ws.Cells(newrow + 27, 12) = itemref + 1
ws.Cells(newrow + 27, 18) = itemref + 2
ws.Cells(newrow + 27, 24) = itemref + 3
newrow = newrow + 26
ws.Rows(newrow + 1).PageBreak = xlPageBreakManual
anotherfour = 0
End If
ws.Cells(newrow + 2, newcolumn) = Left(itemform(a, 1), 4)
ws.Cells(newrow + 3, newcolumn) = Mid(itemform(a, 1), 6, 3)
ws.Cells(newrow + 4, newcolumn) = Left(itemform(a, 10), 3)
ws.Cells(newrow + 2, newcolumn + 2) = arrays(17, itemform(a, 15))
ws.Cells(newrow + 4, newcolumn + 2) = Mid(itemform(a, 10), 5, 999)
ws.Cells(newrow + 5, newcolumn) = "N"
ws.Cells(newrow + 6, newcolumn) = itemform(a, 2)
ws.Cells(newrow + 7, newcolumn) = itemform(a, 3)
ws.Cells(newrow + 8, newcolumn) = arrays(15, itemform(a, 15))
ws.Cells(newrow + 9, newcolumn) = Left(itemform(a, 12), 1)
ws.Cells(newrow + 14, newcolumn) = itemform(a, 4)
ws.Cells(newrow + 15, newcolumn) = itemform(a, 5)
ws.Cells(newrow + 16, newcolumn) = itemform(a, 7)
ws.Cells(newrow + 17, newcolumn) = itemform(a, 8)
ws.Cells(newrow + 18, newcolumn) = itemform(a, 9)
If Left(itemform(a, 1), 4) > 1999 And Left(itemform(a, 1), 4) < 2100 Then ws.Cells(newrow + 19, newcolumn) = itemform(a, 11) * 2 Else If arrays(15, itemform(a, 15)) <> "CI" Then ws.Cells(newrow + 19, newcolumn) = itemform(a, 11)
If arrays(15, itemform(a, 15)) = "US" Then
ws.Cells(newrow + 20, newcolumn) = itemform(a, 11) / 2
ws.Cells(newrow + 21, newcolumn) = itemform(a, 11) / 2
End If
 If Left(itemform(a, 1), 4) > 1999 And Left(itemform(a, 1), 4) < 2100 Then
ws.Cells(newrow + 20, newcolumn) = itemform(a, 11)
ws.Cells(newrow + 21, newcolumn) = itemform(a, 11)
 End If
If arrays(15, itemform(a, 15)) = "US" Then
ws.Cells(newrow + 23, newcolumn + 3) = "100%"
ws.Cells(newrow + 24, newcolumn) = "CI"
ws.Cells(newrow + 24, newcolumn + 3) = itemform(a, 6)
End If
newcolumn = newcolumn + 6
If newcolumn = 18 And wsno = 1 Then newcolumn = 19
If newcolumn = 31 And wsno = 1 Then newcolumn = 37
If newcolumn = 67 Then
wsno = 2
newcolumn = 6
End If
If wsno = 2 And newcolumn = 30 Then
newcolumn = 6
anotherfour = 1
End If
End If
a = a + 1
Loop
aaa = 1
Do While aaa < lastrefill
If setrefill(aaa, 0) <> "" Then
If wsno = 2 And newcolumn = 6 And newrow = 0 And anotherfour = 0 Then
ThisWorkbook.Sheets("Folder3").EnableSelection = xlNoRestrictions
ThisWorkbook.Sheets("Folder3").Visible = True
ThisWorkbook.Sheets("Folder3").Copy
ThisWorkbook.Sheets("Folder3").Unprotect
ThisWorkbook.Sheets("Folder3").Visible = False
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Folder3")
End If
If anotherfour = 1 Then
Dim rgsource2 As Range
Dim rgdest2
Set rgsource2 = ws.Range("A1:A25")
Set rgdest2 = ws.Range("A" & newrow + 27)
rgsource2.EntireRow.Copy (rgdest2)
ws.Range("F" & newrow + 28, "AC" & newrow + 51).ClearContents
ws.Cells(newrow + 27, 6) = itemref
ws.Cells(newrow + 27, 12) = itemref + 1
ws.Cells(newrow + 27, 18) = itemref + 2
ws.Cells(newrow + 27, 24) = itemref + 3
newrow = newrow + 26
ws.Rows(newrow + 1).PageBreak = xlPageBreakManual
anotherfour = 0
End If
ws.Cells(newrow + 2, newcolumn) = Left(setrefill(aaa, 1), 4)
ws.Cells(newrow + 3, newcolumn) = Mid(setrefill(aaa, 1), 6, 3)
ws.Cells(newrow + 4, newcolumn) = Right(setrefill(aaa, 1), 3)
ws.Cells(newrow + 2, newcolumn + 2) = "Refill"
ws.Cells(newrow + 5, newcolumn) = "N"
ws.Cells(newrow + 6, newcolumn) = 0
ws.Cells(newrow + 7, newcolumn) = 0
ws.Cells(newrow + 8, newcolumn) = "US"
ws.Cells(newrow + 9, newcolumn) = Left(setrefill(aaa, 3), 1)
ws.Cells(newrow + 10, newcolumn) = setrefill(aaa, 4)
ws.Cells(newrow + 12, newcolumn) = "D"
ws.Cells(newrow + 25, newcolumn) = setrefill(aaa, 2)
newcolumn = newcolumn + 6
If newcolumn = 18 And wsno = 1 Then newcolumn = 19
If newcolumn = 31 And wsno = 1 Then newcolumn = 37
If newcolumn = 67 Then
wsno = 2
newcolumn = 6
End If
If wsno = 2 And newcolumn = 30 Then
newcolumn = 6
anotherfour = 1
End If
End If
aaa = aaa + 1
Loop
If lastuser > 12 Then
e = 13
ThisWorkbook.Sheets("Folder4").EnableSelection = xlNoRestrictions
ThisWorkbook.Sheets("Folder4").Visible = True
ThisWorkbook.Sheets("Folder4").Copy
ThisWorkbook.Sheets("Folder4").Unprotect
ThisWorkbook.Sheets("Folder4").Visible = False
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Folder4")
userrw = 5
ttop = 26
Do While e < lastuser + 1
If e > ttop Then
Dim rgsource3 As Range
Dim rgdest3
Set rgsource3 = ws.Range("A1:A46")
Set rgdest3 = ws.Range("A" & userrw)
rgsource3.EntireRow.Copy (rgdest3)
ws.Range("A" & userrw + 4, "AG" & userrw + 45).ClearContents
ws.Cells(userrw + 4, 1) = e
ws.Cells(userrw + 7, 1) = e + 1
ws.Cells(userrw + 10, 1) = e + 2
ws.Cells(userrw + 13, 1) = e + 3
ws.Cells(userrw + 16, 1) = e + 4
ws.Cells(userrw + 19, 1) = e + 5
ws.Cells(userrw + 22, 1) = e + 6
ws.Cells(userrw + 25, 1) = e + 7
ws.Cells(userrw + 28, 1) = e + 8
ws.Cells(userrw + 31, 1) = e + 9
ws.Cells(userrw + 34, 1) = e + 10
ws.Cells(userrw + 37, 1) = e + 11
ws.Cells(userrw + 40, 1) = e + 12
ws.Cells(userrw + 43, 1) = e + 13
ttop = e + 13
ws.Rows(userrw).PageBreak = xlPageBreakManual
userrw = userrw + 4
End If
ws.Cells(userrw, 3) = users(e, 0)
ws.Cells(userrw + 1, 3) = users(e, 1)
ws.Cells(userrw, 9) = users(e, 2)
ws.Cells(userrw, 11) = users(e, 3)
ws.Cells(userrw, 13) = users(e, 4)
If users(e, 5) <> "" Then ws.Cells(userrw, 18) = iitem(Left(users(e, 5), 2))
ws.Cells(userrw, 20) = users(e, 8)
ws.Cells(userrw, 21) = users(e, 9)
If ws.Cells(userrw, 18) <> "" Then ws.Cells(userrw, 22) = "A"
ws.Cells(userrw, 24) = users(e, 7)
ws.Cells(userrw, 30) = users(e, 6)
If users(e, 10) = True Then ws.Cells(userrw, 27) = "Y"
If users(e, 11) = True Then ws.Cells(userrw, 28) = "Y"
If users(e, 12) = True Then ws.Cells(userrw, 29) = "Y"
If users(e, 13) <> "" Then ws.Cells(userrw + 1, 18) = iitem(Left(users(e, 13), 2))
ws.Cells(userrw + 1, 20) = users(e, 16)
ws.Cells(userrw + 1, 21) = users(e, 17)
If ws.Cells(userrw + 1, 18) <> "" Then ws.Cells(userrw + 1, 22) = "A"
ws.Cells(userrw + 1, 24) = users(e, 15)
ws.Cells(userrw + 1, 30) = users(e, 14)
If users(e, 18) = True Then ws.Cells(userrw + 1, 27) = "Y"
If users(e, 19) = True Then ws.Cells(userrw + 1, 28) = "Y"
If users(e, 20) = True Then ws.Cells(userrw + 1, 29) = "Y"
If users(e, 21) <> "" Then ws.Cells(userrw + 2, 18) = iitem(Left(users(e, 21), 2))
ws.Cells(userrw + 2, 20) = users(e, 24)
ws.Cells(userrw + 2, 21) = users(e, 25)
If ws.Cells(userrw + 2, 18) <> "" Then ws.Cells(userrw + 2, 22) = "A"
ws.Cells(userrw + 2, 24) = users(e, 23)
ws.Cells(userrw + 2, 30) = users(e, 22)
If users(e, 26) = True Then ws.Cells(userrw + 2, 27) = "Y"
If users(e, 27) = True Then ws.Cells(userrw + 2, 28) = "Y"
If users(e, 28) = True Then ws.Cells(userrw + 2, 29) = "Y"
e = e + 1
userrw = userrw + 3
Loop
End If
End Sub
Private Sub CommandButton9_Click()
a = Label96.Caption
If a <> 1 Then
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
a = a - 1
Label96.Caption = a
TextBox70.Text = users(a, 0)
TextBox71.Text = users(a, 1)
TextBox72.Text = users(a, 2)
TextBox73.Text = users(a, 3)
TextBox74.Text = users(a, 4)
ComboBox20.Value = users(a, 5)
TextBox58.Text = users(a, 6)
TextBox61.Text = users(a, 7)
TextBox64.Text = users(a, 8)
TextBox65.Text = users(a, 9)
CheckBox1.Value = users(a, 10)
CheckBox4.Value = users(a, 11)
CheckBox7.Value = users(a, 12)
ComboBox21.Value = users(a, 13)
TextBox59.Text = users(a, 14)
TextBox62.Text = users(a, 15)
TextBox66.Text = users(a, 16)
TextBox67.Text = users(a, 17)
CheckBox2.Value = users(a, 18)
CheckBox5.Value = users(a, 19)
CheckBox8.Value = users(a, 20)
ComboBox22.Value = users(a, 21)
TextBox60.Text = users(a, 22)
TextBox63.Text = users(a, 23)
TextBox68.Text = users(a, 24)
TextBox69.Text = users(a, 25)
CheckBox3.Value = users(a, 26)
CheckBox6.Value = users(a, 27)
CheckBox9.Value = users(a, 28)
End If
End Sub
Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
MsgBox ("Made by Andrew Aaron - aaaron@gkservices.com - last updated February 18, 2015")
End Sub

Private Sub MultiPage1_Change()
If MultiPage1.Value = 1 Then
ComboBox20.Clear
ComboBox21.Clear
ComboBox22.Clear
ComboBox20.AddItem
ComboBox21.AddItem
ComboBox22.AddItem
a = 1
b = 0
Do While a < lastitem + 1
If itemform(a, 13) = True And itemform(a, 0) <> "" Then
If arrays(13, itemform(a, 15)) = "Garment" Then
b = b + 1
ComboBox20.AddItem a & " " & itemform(a, 0)
ComboBox21.AddItem a & " " & itemform(a, 0)
ComboBox22.AddItem a & " " & itemform(a, 0)
usergarment(b) = a
End If
End If
a = a + 1
Loop
If lastuser = "" Then lastuser = 1
ComboBox20.Value = users(Label96.Caption, 5)
ComboBox21.Value = users(Label96.Caption, 13)
ComboBox22.Value = users(Label96.Caption, 21)
Else
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
bb = 1
Do While bb < lastitem + 1
itemform(bb, 16) = ""
bb = bb + 1
Loop
aa = 1
Do While aa < lastuser + 1
If users(aa, 29) > 0 Then itemform(usergarment(users(aa, 29)), 16) = "Used"
If users(aa, 30) > 0 Then itemform(usergarment(users(aa, 30)), 16) = "Used"
If users(aa, 31) > 0 Then itemform(usergarment(users(aa, 31)), 16) = "Used"
aa = aa + 1
Loop
End If
If MultiPage1.Value = 2 Then
TextBox106.Text = "$" & TextBox11.Text & "/" & TextBox20.Text & "%"
TextBox107.Text = "$" & TextBox13.Text & "/" & TextBox21.Text & "%"
TextBox108.Text = "$" & Round(TextBox9.Text, 0)
If TextBox109.Text = "" Or TextBox109.Text = "   " Then TextBox109.Text = TextBox26.Text & " " & TextBox27.Text & " " & TextBox114.Text & " " & TextBox28.Text
If TextBox110.Text = "" Then TextBox110.Text = TextBox29.Text
If ComboBox18.Value = "None" Then ip = "NO" Else ip = "YES"
If ComboBox17.Value = "None" Then ip = ip & "/NO" Else ip = ip & "/YES"
TextBox111.Text = ip
End If
End Sub
Private Sub OptionButton4_Click()
If scrolling <> 1 Then itemform(showing, 13) = OptionButton4.Value: Call calculate
End Sub
Private Sub OptionButton6_Click()
If scrolling <> 1 Then itemform(showing + 1, 13) = OptionButton6.Value: Call calculate
End Sub
Private Sub OptionButton8_Click()
If scrolling <> 1 Then itemform(showing + 2, 13) = OptionButton8.Value: Call calculate
End Sub
Private Sub OptionButton3_Click()
If scrolling <> 1 Then
If itemform(showing, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
OptionButton4.Value = True
Exit Sub
End If
itemform(showing, 13) = OptionButton4.Value
Call calculate
End If
End Sub
Private Sub OptionButton5_Click()
If scrolling <> 1 Then
If itemform(showing + 1, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
OptionButton6.Value = True
Exit Sub
End If
itemform(showing + 1, 13) = OptionButton6.Value
Call calculate
End If
End Sub
Private Sub OptionButton7_Click()
If scrolling <> 1 Then
If itemform(showing + 2, 16) = "Used" Then
MsgBox "Sorry, you can't change an item when it's assigned to a wearer in the Folder section"
OptionButton8.Value = True
Exit Sub
End If
itemform(showing + 2, 13) = OptionButton8.Value
Call calculate
End If
End Sub
Private Sub ScrollBar2_Change()
scrolling = 1
Label13 = "#" & ScrollBar2.Value + 1
Label31 = "#" & ScrollBar2.Value + 2
Label43 = "#" & ScrollBar2.Value + 3
showing = ScrollBar2.Value + 1
If ScrollBar2.Value < 96 And ScrollBar2.Max = ScrollBar2.Value Then ScrollBar2.Max = ScrollBar2.Value + 1
ComboBox6.Value = valueofitem(1)
ComboBox5.Clear
a = 0
Line = 0
Do While a < total
If ComboBox6.ListIndex = 0 Or ComboBox6.Text = arrays(19, a) Then ComboBox5.AddItem arrays(1, a): list1(Line) = a: Line = Line + 1
a = a + 1
Loop
If itemform(showing, 15) <> "" Then ComboBox5.Value = arrays(1, itemform(showing, 15)) Else ComboBox5.Value = ""
TextBox3.Text = itemform(showing, 1)
TextBox83.Text = itemform(showing, 2)
C = itemform(showing, 15)
ComboBox4.Clear
'MsgBox c
Dim colour() As String
If itemform(showing, 0) <> "" Then
If InStr(arrays(18, C), ",") = 0 Then
ComboBox4.AddItem arrays(18, C)
ComboBox4.ListIndex = 0
TextBox3.Text = Left(TextBox3.Text, 8) & "-" & Left(ComboBox4.Value, 3)
Else
colour = Split(arrays(18, C), ",")
Dim i As Integer
For i = LBound(colour) To UBound(colour)
ComboBox4.AddItem colour(i)
Next i
End If
End If
ComboBox4.Value = itemform(showing, 10)
ComboBox7.Value = valueofitem2(1)
If itemform(showing, 13) = False Or itemform(showing, 13) = "" Then OptionButton3.Value = True Else OptionButton4.Value = True
TextBox2.Text = itemform(showing, 11)
If ComboBox5.ListIndex >= 0 Then TextBox82.Text = arrays(2, itemform(showing, 15)) Else TextBox82.Text = ""
ComboBox11.Value = valueofitem(2)
ComboBox10.Clear
Line = 0
a = 0
Do While a < total
If ComboBox11.ListIndex = 0 Or ComboBox11.Text = arrays(19, a) Then ComboBox10.AddItem arrays(1, a): list2(Line) = a: Line = Line + 1
a = a + 1
Loop
If itemform(showing + 1, 15) <> "" Then ComboBox10.Value = arrays(1, itemform(showing + 1, 15)) Else ComboBox10.Value = ""
TextBox4.Text = itemform(showing + 1, 1)
TextBox86.Text = itemform(showing + 1, 2)
C = itemform(showing + 1, 15)
ComboBox9.Clear
'MsgBox c
If itemform(showing + 1, 0) <> "" Then
If InStr(arrays(18, C), ",") = 0 Then
ComboBox9.AddItem arrays(18, C)
ComboBox9.ListIndex = 0
TextBox4.Text = Left(TextBox4.Text, 8) & "-" & Left(ComboBox9.Value, 3)
Else
colour2 = Split(arrays(18, C), ",")
For i = LBound(colour2) To UBound(colour2)
ComboBox9.AddItem colour2(i)
Next i
End If
End If
ComboBox9.Value = itemform(showing + 1, 10)
ComboBox8.Value = valueofitem2(2)
If itemform(showing + 1, 13) = False Or itemform(showing + 1, 13) = "" Then OptionButton5.Value = True Else OptionButton6.Value = True
TextBox5.Text = itemform(showing + 1, 11)
If ComboBox10.ListIndex >= 0 Then TextBox85.Text = arrays(2, itemform(showing + 1, 15)) Else TextBox85.Text = ""
ComboBox15.Value = valueofitem(3)
ComboBox14.Clear
Line = 0
a = 0
Do While a < total
If ComboBox15.ListIndex = 0 Or ComboBox15.Text = arrays(19, a) Then ComboBox14.AddItem arrays(1, a): list3(Line) = a: Line = Line + 1
a = a + 1
Loop
If itemform(showing + 2, 15) <> "" Then ComboBox14.Value = arrays(1, itemform(showing + 2, 15)) Else ComboBox14.Value = ""
TextBox6.Text = itemform(showing + 2, 1)
TextBox89.Text = itemform(showing + 2, 2)
C = itemform(showing + 2, 15)
ComboBox13.Clear
'MsgBox c
If itemform(showing + 2, 0) <> "" Then
If InStr(arrays(18, C), ",") = 0 Then
ComboBox13.AddItem arrays(18, C)
ComboBox13.ListIndex = 0
TextBox6.Text = Left(TextBox6.Text, 8) & "-" & Left(ComboBox13.Value, 3)
Else
colour3 = Split(arrays(18, C), ",")
For i = LBound(colour3) To UBound(colour3)
ComboBox13.AddItem colour3(i)
Next i
End If
End If
ComboBox13.Value = itemform(showing + 2, 10)
ComboBox12.Value = valueofitem2(3)
If itemform(showing + 2, 13) = False Or itemform(showing + 2, 13) = "" Then OptionButton7.Value = True Else OptionButton8.Value = True
TextBox7.Text = itemform(showing + 2, 11)
If ComboBox14.ListIndex >= 0 Then TextBox90.Text = arrays(2, itemform(showing + 2, 15)) Else TextBox90.Text = ""
If TextBox6.Text <> "" And ComboBox13.Value <> "" Then TextBox6.Text = Left(TextBox6.Text, 8) & "-" & Left(ComboBox13.Value, 3)
If TextBox3.Text <> "" And ComboBox4.Value <> "" Then TextBox3.Text = Left(TextBox3.Text, 8) & "-" & Left(ComboBox4.Value, 3)
If TextBox4.Text <> "" And ComboBox9.Value <> "" Then TextBox4.Text = Left(TextBox4.Text, 8) & "-" & Left(ComboBox9.Value, 3)
scrolling = 0
End Sub
Private Sub ScrollBar4_Change()
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
Label113 = ScrollBar4.Value + 1
Label114 = ScrollBar4.Value + 2
Label115 = ScrollBar4.Value + 3
If setrefill(Label113, 0) <> "" Then ComboBox23.Value = setrefill(Label113, 0) Else ComboBox23.ListIndex = 0
TextBox75.Text = setrefill(Label113, 1)
TextBox78.Value = setrefill(Label113, 2)
If setrefill(Label113, 3) <> "" Then ComboBox27.Value = setrefill(Label113, 3)
If setrefill(Label113, 4) <> "" Then ComboBox30.Value = setrefill(Label113, 4)
If setrefill(Label114, 0) <> "" Then ComboBox24.Value = setrefill(Label114, 0) Else ComboBox24.ListIndex = 0
TextBox76.Text = setrefill(Label114, 1)
TextBox79.Value = setrefill(Label114, 2)
If setrefill(Label114, 3) <> "" Then ComboBox28.Value = setrefill(Label114, 3)
If setrefill(Label114, 4) <> "" Then ComboBox31.Value = setrefill(Label114, 4)
If setrefill(Label115, 0) <> "" Then ComboBox25.Value = setrefill(Label115, 0) Else ComboBox25.ListIndex = 0
TextBox77.Text = setrefill(Label115, 1)
TextBox80.Value = setrefill(Label115, 2)
If setrefill(Label115, 3) <> "" Then ComboBox29.Value = setrefill(Label115, 3)
If setrefill(Label115, 4) <> "" Then ComboBox32.Value = setrefill(Label115, 4)
If ScrollBar4.Value = ScrollBar4.Max And ScrollBar4.Value < 96 Then ScrollBar4.Max = ScrollBar4.Max + 1
If lastrefill < ScrollBar4.Value + 3 Then lastrefill = ScrollBar4.Value + 3
End Sub
Private Sub TextBox11_Change()
Call calculate
End Sub
Private Sub TextBox117_Change()
Call calculate
End Sub
Private Sub TextBox13_Change()
Call calculate
End Sub
Private Sub TextBox2_Change()
itemform(showing, 11) = TextBox2.Text
Call calculate
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub
Private Sub TextBox20_Change()
Call calculate
End Sub
Private Sub TextBox21_Change()
Call calculate
End Sub
Private Sub TextBox5_Change()
itemform(showing + 1, 11) = TextBox5.Text
Call calculate
End Sub
Private Sub TextBox7_Change()
itemform(showing + 2, 11) = TextBox7.Text
Call calculate
End Sub
Private Sub TextBox58_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox59_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox60_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox61_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox62_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox63_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox64_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox65_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox66_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox67_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox68_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox69_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox70_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox71_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox72_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox73_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub
Private Sub TextBox74_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton10_Click: KeyCode = 0
End Sub

Private Sub TextBox82_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UserForm3.TextBox1.Text = TextBox82.Text
UserForm3.Show
End Sub
Sub calculate()
On Error Resume Next
Dim a, b, C, d, e, f, g, h, i, j, k, l, m, n, imagetotal, preptotal As Double
aa = 1
Subtotal = 0
imagetotal = 0
preptotal = 0
Do While aa < lastitem + 1
If itemform(aa, 11) = "" Then a = 0 Else a = itemform(aa, 11)
If itemform(aa, 13) <> True Or itemform(aa, 2) = "" Then b = 0 Else b = itemform(aa, 2)
Subtotal = Subtotal + (a * b) + (itemform(aa, 3) * Round(a * itemform(aa, 6), 0))
If ComboBox17.Value <> "None" Then
preptotal = preptotal + (WorksheetFunction.VLookup(arrays(24, itemform(aa, 15)) & ComboBox17.Value, Worksheets("ImagePrep").Range("C:D"), 2, False) * itemform(aa, 11))
End If
If ComboBox18.Value <> "None" Then
imagetotal = imagetotal + (WorksheetFunction.VLookup(arrays(25, itemform(aa, 15)) & ComboBox18.Value, Worksheets("ImagePrep").Range("A:B"), 2, False) * itemform(aa, 11))
End If
aa = aa + 1
Loop
If TextBox2.Value = "" Then a = 0 Else a = TextBox2.Text
If TextBox83.Value = "" Then b = 0 Else b = TextBox83.Text
If TextBox5.Value = "" Then C = 0 Else C = TextBox5.Text
If TextBox86.Value = "" Then d = 0 Else d = TextBox86.Text
If TextBox7.Value = "" Then e = 0 Else e = TextBox7.Text
If TextBox89.Value = "" Then f = 0 Else f = TextBox89.Text
If OptionButton4.Value = True Then TextBox84.Text = Format((a * b) + (itemform(showing, 3) * Round(a * itemform(showing, 6), 0)), "Currency") Else TextBox84.Text = 0
If OptionButton6.Value = True Then TextBox87.Text = Format((C * d) + (itemform(showing + 1, 3) * Round(C * itemform(showing + 1, 6), 0)), "Currency") Else TextBox87.Text = 0
If OptionButton8.Value = True Then TextBox88.Text = Format((e * f) + (itemform(showing + 2, 3) * Round(e * itemform(showing + 2, 6), 0)), "Currency") Else TextBox88.Text = 0
If Subtotal < CDbl(TextBox9.Text) Then TextBox10.Text = Format(TextBox9.Text - Subtotal, "currency") Else TextBox10.Text = Format(0, "Currency")
If TextBox20.Text = "" Then n = 0 Else n = TextBox20.Text
m = (Subtotal + TextBox10.Text) * (n / 100) * 1
l = TextBox11.Text * 1
If l > m Then TextBox12.Text = Format(l, "Currency") Else TextBox12.Text = Format(m, "Currency")
If TextBox21.Text = "" Then n = 0 Else n = TextBox21.Text
m = Format((Subtotal + TextBox10.Text) * (n / 100), "Currency") * 1
l = CDbl(TextBox13.Text) * 1
If l > m Then TextBox14.Text = Format(l, "Currency") Else TextBox14.Text = Format(m, "Currency")
TextBox16.Text = Format(preptotal, "Currency")
TextBox15.Text = Format(imagetotal, "Currency")
g = Subtotal + CDbl(TextBox12.Text) + CDbl(TextBox14.Text) + CDbl(TextBox10.Text) + CDbl(TextBox16.Text) + CDbl(TextBox15.Text)
If IsNumeric(TextBox117.Text) = True And TextBox117.Text > 0 Then tax = TextBox117.Text / 100 Else tax = 0
h = Round(g * tax, 2)
i = Round(g + h, 2)
TextBox17.Text = Format(g, "Currency")
TextBox18.Text = Format(h, "Currency")
TextBox19.Text = Format(i, "Currency")
End Sub
Private Sub TextBox83_Change()
Call calculate
End Sub
Private Sub TextBox85_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UserForm3.TextBox1.Text = TextBox85.Text
UserForm3.Show
End Sub
Private Sub TextBox86_Change()
Call calculate
End Sub
Private Sub TextBox89_Change()
Call calculate
End Sub
Private Sub TextBox9_Change()
Call calculate
End Sub
Private Sub TextBox90_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UserForm3.TextBox1.Text = TextBox90.Text
UserForm3.Show
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
If KeyAscii = 46 And InStr(TextBox9.Text, ".") > 0 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        If KeyAscii <> 46 Then KeyAscii = 0: Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
If KeyAscii = 46 And InStr(TextBox11.Text, ".") > 0 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        If KeyAscii <> 46 Then KeyAscii = 0: Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
If KeyAscii = 46 And InStr(TextBox13.Text, ".") > 0 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        If KeyAscii <> 46 Then KeyAscii = 0: Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox20_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
If KeyAscii = 46 And InStr(TextBox20.Text, ".") > 0 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        If KeyAscii <> 46 Then KeyAscii = 0: Beep
    End Select
End Sub
'~~> Disable Pasting CTRL V , SHIFT + INSERT
Private Sub TextBox21_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert) Then
        KeyCode = 0
    End If
End Sub
'~~> Preventing input of non numerics
Private Sub TextBox21_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 36 And KeyAscii < 41 Then KeyAscii = 0: Beep
If KeyAscii = 46 And InStr(TextBox21.Text, ".") > 0 Then KeyAscii = 0: Beep
    Select Case KeyAscii
      Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyLeft, _
      vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
      Case Else
        If KeyAscii <> 46 Then KeyAscii = 0: Beep
    End Select
End Sub
Private Sub UserForm_Initialize()
With ComboBox16
    .AddItem "Small ($35-100)"
    .AddItem "Medium ($101-250)"
    .AddItem "Large ($250+)"
    .ListIndex = 0
End With
With ComboBox19
    .AddItem "Yes"
    .AddItem "No"
    .ListIndex = 1
End With
With ComboBox26
    .AddItem "Yes"
    .AddItem "No"
    .ListIndex = 1
End With
With ComboBox18
.AddItem "None"
.AddItem "B"
.AddItem "B2"
.AddItem "S1"
.AddItem "S"
.AddItem "G1"
.AddItem "G"
.AddItem "P1"
.AddItem "P"
.AddItem "P2"
.AddItem "P3"
.AddItem "P4"
.AddItem "P5"
.AddItem "P6"
.ListIndex = 0
End With
With ComboBox17
.AddItem "None"
.AddItem "Chrome"
.AddItem "Copper"
.AddItem "Bronze"
.AddItem "Jade"
.AddItem "Silver"
.AddItem "Gold"
.AddItem "Platinum"
.AddItem "Diamond"
.ListIndex = 0
End With
With ComboBox27
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 2
End With
With ComboBox28
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 2
End With
With ComboBox29
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 2
End With
With ComboBox30
.AddItem "A"
.AddItem "B"
.AddItem "C"
.AddItem "D"
.ListIndex = 0
End With
With ComboBox31
.AddItem "A"
.AddItem "B"
.AddItem "C"
.AddItem "D"
.ListIndex = 0
End With
With ComboBox32
.AddItem "A"
.AddItem "B"
.AddItem "C"
.AddItem "D"
.ListIndex = 0
End With
With ComboBox7
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 0
End With
With ComboBox8
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 0
End With
With ComboBox12
.AddItem "WEEKLY"
.AddItem "BI-WEEKLY"
.AddItem "MONTHLY"
.ListIndex = 0
End With
dp = 1
DPCs = Worksheets("DPCs").UsedRange.Rows.Count
Do While dp <= DPCs
ComboBox33.AddItem (Sheets("DPCs").Range("A" & dp).Value & "-" & Sheets("DPCs").Range("B" & dp).Value)
dp = dp + 1
Loop
    ComboBox6.AddItem "ALL"
    ComboBox11.AddItem "ALL"
    ComboBox15.AddItem "ALL"
    ComboBox6.ListIndex = 0
    ComboBox11.ListIndex = 0
    ComboBox15.ListIndex = 0
 Const ProductSheetName = "ItemList"
 b = Worksheets("ItemList").UsedRange.Rows.Count
  ProductRange = "S2:S" & b
  Const ResultsCol = "E"
  Dim productWS As Worksheet
  Dim uniqueList() As String
  Dim productsList As Range
  Dim anyProduct
  Dim LC As Integer
 ReDim uniqueList(1 To 1)
  Set productWS = Worksheets(ProductSheetName)
  Set productsList = productWS.Range(ProductRange)
  Application.ScreenUpdating = False
  For Each anyProduct In productsList
        For LC = LBound(uniqueList) To UBound(uniqueList)
          If Trim(anyProduct) = uniqueList(LC) Then
            Exit For ' found match, exit
          End If
        Next
        If LC > UBound(uniqueList) Then
          'new item, add it
          uniqueList(UBound(uniqueList)) = Trim(anyProduct)
          'make room for another
          ReDim Preserve uniqueList(1 To UBound(uniqueList) + 1)
    End If
  Next ' end anyProduct loop
  If UBound(uniqueList) > 1 Then
    'remove empty element
    ReDim Preserve uniqueList(1 To UBound(uniqueList) - 1)
  End If
  For LC = LBound(uniqueList) To UBound(uniqueList)
    ComboBox6.AddItem uniqueList(LC)
    ComboBox11.AddItem uniqueList(LC)
    ComboBox15.AddItem uniqueList(LC)
  Next
  'housekeeping cleanup
  Set productsList = Nothing
  Set productWS = Nothing
total = 0
For i = 0 To b
arrays(1, i) = Sheets("ItemList").Range("A" & i + 2)
arrays(2, i) = Sheets("ItemList").Range("B" & i + 2)
arrays(3, i) = Format(Sheets("ItemList").Range("C" & i + 2), "fixed")
arrays(4, i) = Format(Sheets("ItemList").Range("D" & i + 2), "fixed")
arrays(5, i) = Format(Sheets("ItemList").Range("E" & i + 2), "fixed")
arrays(6, i) = Format(Sheets("ItemList").Range("F" & i + 2), "fixed")
arrays(7, i) = Format(Sheets("ItemList").Range("G" & i + 2), "fixed")
arrays(8, i) = Format(Sheets("ItemList").Range("H" & i + 2), "fixed")
arrays(9, i) = Format(Sheets("ItemList").Range("I" & i + 2), "fixed")
arrays(10, i) = Format(Sheets("ItemList").Range("J" & i + 2), "fixed")
arrays(11, i) = Format(Sheets("ItemList").Range("K" & i + 2), "fixed")
arrays(12, i) = Sheets("ItemList").Range("L" & i + 2)
arrays(13, i) = Sheets("ItemList").Range("M" & i + 2)
arrays(14, i) = Sheets("ItemList").Range("N" & i + 2)
arrays(15, i) = Sheets("ItemList").Range("O" & i + 2)
arrays(16, i) = Format(Sheets("ItemList").Range("P" & i + 2), "fixed")
arrays(17, i) = Sheets("ItemList").Range("Q" & i + 2)
arrays(18, i) = Sheets("ItemList").Range("R" & i + 2)
arrays(19, i) = Sheets("ItemList").Range("S" & i + 2)
arrays(20, i) = Sheets("ItemList").Range("T" & i + 2)
arrays(21, i) = Format(Sheets("ItemList").Range("U" & i + 2), "fixed")
arrays(22, i) = Format(Sheets("ItemList").Range("V" & i + 2), "fixed")
arrays(23, i) = Format(Sheets("ItemList").Range("W" & i + 2), "fixed")
arrays(24, i) = Sheets("ItemList").Range("X" & i + 2)
arrays(25, i) = Sheets("ItemList").Range("Y" & i + 2)
total = total + 1
Next
bb = Worksheets("Refills").UsedRange.Rows.Count
ComboBox23.AddItem
ComboBox24.AddItem
ComboBox25.AddItem
For ii = 0 To bb - 2
refills(ii, 0) = Sheets("Refills").Range("A" & ii + 2)
refills(ii, 1) = Sheets("Refills").Range("B" & ii + 2)
ComboBox23.AddItem refills(ii, 0)
ComboBox24.AddItem refills(ii, 0)
ComboBox25.AddItem refills(ii, 0)
Next
total = total - 2
showing = 1
Call filter(1)
Call filter(2)
Call filter(3)
If taxrate = "" Then TextBox117.Text = Sheets("Data4").Range("A71").Value Else TextBox117.Text = taxrate
If salesrep = "" Then TextBox112.Text = Sheets("Data4").Range("A72").Value Else TextBox112.Text = salesrep
If dpc = "" Then ComboBox33.Value = Sheets("Data4").Range("A73").Value Else ComboBox33.Value = dpc
If Sheets("Data4").Range("A9").Value <> "" Then
For p = 2 To 99
For j = 1 To 16
itemform(p - 1, j - 1) = Sheets("Data1").Cells(p, j).Value
Next
Next
For p = 2 To 99
For j = 1 To 5
setrefill(p - 1, j - 1) = Sheets("Data2").Cells(p, j).Value
Next
Next
For p = 2 To 999
For j = 1 To 32
users(p - 1, j - 1) = Sheets("Data3").Cells(p, j).Value
Next
Next
TextBox9.Text = Sheets("Data4").Range("A1").Value
TextBox11.Text = Sheets("Data4").Range("A2").Value
TextBox20.Text = Sheets("Data4").Range("A3").Value
TextBox13.Text = Sheets("Data4").Range("A4").Value
TextBox21.Text = Sheets("Data4").Range("A5").Value
ComboBox17.Value = Sheets("Data4").Range("A6").Value
ComboBox18.Value = Sheets("Data4").Range("A7").Value
TextBox8.Text = Sheets("Data4").Range("A8").Value
ComboBox16.Value = Sheets("Data4").Range("A9").Value
TextBox22.Text = Sheets("Data4").Range("A10").Value
TextBox23.Text = Sheets("Data4").Range("A11").Value
TextBox24.Text = Sheets("Data4").Range("A12").Value
TextBox25.Text = Sheets("Data4").Range("A13").Value
TextBox26.Text = Sheets("Data4").Range("A14").Value
TextBox27.Text = Sheets("Data4").Range("A15").Value
TextBox28.Text = Sheets("Data4").Range("A16").Value
TextBox29.Text = Sheets("Data4").Range("A17").Value
TextBox30.Text = Sheets("Data4").Range("A18").Value
TextBox31.Text = Sheets("Data4").Range("A19").Value
TextBox32.Text = Sheets("Data4").Range("A20").Value
TextBox33.Text = Sheets("Data4").Range("A21").Value
TextBox34.Text = Sheets("Data4").Range("A22").Value
TextBox35.Text = Sheets("Data4").Range("A23").Value
TextBox36.Text = Sheets("Data4").Range("A24").Value
TextBox37.Text = Sheets("Data4").Range("A25").Value
TextBox38.Text = Sheets("Data4").Range("A26").Value
TextBox39.Text = Sheets("Data4").Range("A27").Value
TextBox40.Text = Sheets("Data4").Range("A28").Value
TextBox41.Text = Sheets("Data4").Range("A29").Value
TextBox42.Text = Sheets("Data4").Range("A30").Value
TextBox43.Text = Sheets("Data4").Range("A31").Value
TextBox44.Text = Sheets("Data4").Range("A32").Value
TextBox45.Text = Sheets("Data4").Range("A33").Value
TextBox46.Text = Sheets("Data4").Range("A34").Value
TextBox47.Text = Sheets("Data4").Range("A35").Value
TextBox48.Text = Sheets("Data4").Range("A36").Value
TextBox49.Text = Sheets("Data4").Range("A37").Value
TextBox50.Text = Sheets("Data4").Range("A38").Value
TextBox51.Text = Sheets("Data4").Range("A39").Value
TextBox52.Text = Sheets("Data4").Range("A40").Value
TextBox53.Text = Sheets("Data4").Range("A41").Value
TextBox55.Text = Sheets("Data4").Range("A42").Value
TextBox56.Text = Sheets("Data4").Range("A43").Value
TextBox57.Text = Sheets("Data4").Range("A44").Value
ComboBox19.Value = Sheets("Data4").Range("A45").Value
ComboBox26.Value = Sheets("Data4").Range("A46").Value
TextBox81.Text = Sheets("Data4").Range("A47").Value
TextBox91.Text = Sheets("Data4").Range("A48").Value
lastuser = Sheets("Data4").Range("A49").Value
lastitem = Sheets("Data4").Range("A50").Value
lastrefill = Sheets("Data4").Range("A51").Value
TextBox114.Text = Sheets("Data4").Range("A52").Value
TextBox115.Text = Sheets("Data4").Range("A53").Value
TextBox92.Text = Sheets("Data4").Range("A54").Value
TextBox93.Text = Sheets("Data4").Range("A55").Value
TextBox94.Text = Sheets("Data4").Range("A56").Value
TextBox95.Text = Sheets("Data4").Range("A57").Value
TextBox96.Text = Sheets("Data4").Range("A58").Value
TextBox97.Text = Sheets("Data4").Range("A59").Value
TextBox98.Text = Sheets("Data4").Range("A60").Value
TextBox99.Text = Sheets("Data4").Range("A61").Value
TextBox100.Text = Sheets("Data4").Range("A62").Value
TextBox101.Text = Sheets("Data4").Range("A63").Value
TextBox102.Text = Sheets("Data4").Range("A64").Value
TextBox103.Text = Sheets("Data4").Range("A65").Value
TextBox104.Text = Sheets("Data4").Range("A66").Value
TextBox105.Text = Sheets("Data4").Range("A67").Value
TextBox106.Text = Sheets("Data4").Range("A68").Value
TextBox107.Text = Sheets("Data4").Range("A69").Value
TextBox108.Text = Sheets("Data4").Range("A70").Value
TextBox118.Text = Sheets("Data4").Range("A74").Value
TextBox119.Text = Sheets("Data4").Range("A75").Value
ScrollBar2.Value = 0
Call ScrollBar2_Change
If setrefill(Label113, 0) <> "" Then ComboBox23.Value = setrefill(Label113, 0) Else ComboBox23.ListIndex = 0
TextBox75.Text = setrefill(Label113, 1)
TextBox78.Value = setrefill(Label113, 2)
If setrefill(Label113, 3) <> "" Then ComboBox27.Value = setrefill(Label113, 3)
If setrefill(Label113, 4) <> "" Then ComboBox30.Value = setrefill(Label113, 4)
If setrefill(Label114, 0) <> "" Then ComboBox24.Value = setrefill(Label114, 0) Else ComboBox24.ListIndex = 0
TextBox76.Text = setrefill(Label114, 1)
TextBox79.Value = setrefill(Label114, 2)
If setrefill(Label114, 3) <> "" Then ComboBox28.Value = setrefill(Label114, 3)
If setrefill(Label114, 4) <> "" Then ComboBox31.Value = setrefill(Label114, 4)
If setrefill(Label115, 0) <> "" Then ComboBox25.Value = setrefill(Label115, 0) Else ComboBox25.ListIndex = 0
TextBox77.Text = setrefill(Label115, 1)
TextBox80.Value = setrefill(Label115, 2)
If setrefill(Label115, 3) <> "" Then ComboBox29.Value = setrefill(Label115, 3)
If setrefill(Label115, 4) <> "" Then ComboBox32.Value = setrefill(Label115, 4)
a = 1
TextBox70.Text = users(a, 0)
TextBox71.Text = users(a, 1)
TextBox72.Text = users(a, 2)
TextBox73.Text = users(a, 3)
TextBox74.Text = users(a, 4)
TextBox58.Text = users(a, 6)
TextBox61.Text = users(a, 7)
TextBox64.Text = users(a, 8)
TextBox65.Text = users(a, 9)
CheckBox1.Value = users(a, 10)
CheckBox4.Value = users(a, 11)
CheckBox7.Value = users(a, 12)
TextBox59.Text = users(a, 14)
TextBox62.Text = users(a, 15)
TextBox66.Text = users(a, 16)
TextBox67.Text = users(a, 17)
CheckBox2.Value = users(a, 18)
CheckBox5.Value = users(a, 19)
CheckBox8.Value = users(a, 20)
TextBox60.Text = users(a, 22)
TextBox63.Text = users(a, 23)
TextBox68.Text = users(a, 24)
TextBox69.Text = users(a, 25)
CheckBox3.Value = users(a, 26)
CheckBox6.Value = users(a, 27)
CheckBox9.Value = users(a, 28)
Call calculate
End If
loaded = 1
End Sub
Public Sub filter(which As Integer)
If scrolling <> 1 Then
If which = 1 Then
ComboBox4.Clear
ComboBox5.Clear
TextBox3.Text = ""
TextBox82.Text = ""
TextBox83.Text = ""
a = 0
Line = 0
Do While a < total
If ComboBox6.ListIndex = 0 Or ComboBox6.Text = arrays(19, a) Then ComboBox5.AddItem arrays(1, a): list1(Line) = a: Line = Line + 1
a = a + 1
Loop
itemform(showing, 14) = ComboBox6.Value
itemform(showing, 0) = ""
itemform(showing, 1) = ""
itemform(showing, 2) = ""
itemform(showing, 3) = ""
itemform(showing, 4) = ""
itemform(showing, 5) = ""
itemform(showing, 6) = ""
itemform(showing, 7) = ""
itemform(showing, 8) = ""
itemform(showing, 9) = ""
itemform(showing, 10) = ""
itemform(showing, 15) = ""
End If
If which = 2 Then
ComboBox9.Clear
ComboBox10.Clear
TextBox4.Text = ""
TextBox85.Text = ""
TextBox86.Text = ""
Line = 0
Do While a < total
If ComboBox11.ListIndex = 0 Or ComboBox11.Text = arrays(19, a) Then ComboBox10.AddItem arrays(1, a): list2(Line) = a: Line = Line + 1
a = a + 1
Loop
itemform(showing + 1, 14) = ComboBox11.Value
itemform(showing + 1, 0) = ""
itemform(showing + 1, 1) = ""
itemform(showing + 1, 2) = ""
itemform(showing + 1, 3) = ""
itemform(showing + 1, 4) = ""
itemform(showing + 1, 5) = ""
itemform(showing + 1, 6) = ""
itemform(showing + 1, 7) = ""
itemform(showing + 1, 8) = ""
itemform(showing + 1, 9) = ""
itemform(showing + 1, 10) = ""
itemform(showing + 1, 15) = ""
End If
If which = 3 Then
ComboBox13.Clear
ComboBox14.Clear
TextBox6.Text = ""
TextBox89.Text = ""
TextBox90.Text = ""
Line = 0
Do While a < total
If ComboBox15.ListIndex = 0 Or ComboBox15.Text = arrays(19, a) Then ComboBox14.AddItem arrays(1, a): list3(Line) = a: Line = Line + 1
a = a + 1
Loop
itemform(showing + 2, 14) = ComboBox15.Value
itemform(showing + 2, 0) = ""
itemform(showing + 2, 1) = ""
itemform(showing + 2, 2) = ""
itemform(showing + 2, 3) = ""
itemform(showing + 2, 4) = ""
itemform(showing + 2, 5) = ""
itemform(showing + 2, 6) = ""
itemform(showing + 2, 7) = ""
itemform(showing + 2, 8) = ""
itemform(showing + 2, 9) = ""
itemform(showing + 2, 10) = ""
itemform(showing + 2, 15) = ""
End If
End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    UserForm1.Hide
  End If
setrefill(Label113, 0) = ComboBox23.Value
setrefill(Label113, 1) = TextBox75.Text
setrefill(Label113, 2) = TextBox78.Value
setrefill(Label113, 3) = ComboBox27.Value
setrefill(Label113, 4) = ComboBox30.Value
setrefill(Label114, 0) = ComboBox24.Value
setrefill(Label114, 1) = TextBox76.Text
setrefill(Label114, 2) = TextBox79.Value
setrefill(Label114, 3) = ComboBox28.Value
setrefill(Label114, 4) = ComboBox31.Value
setrefill(Label115, 0) = ComboBox25.Value
setrefill(Label115, 1) = TextBox77.Text
setrefill(Label115, 2) = TextBox80.Value
setrefill(Label115, 3) = ComboBox29.Value
setrefill(Label115, 4) = ComboBox32.Value
If lastrefill = "" Then lastrefill = 3
a = Label96.Caption
users(a, 0) = TextBox70.Text
users(a, 1) = TextBox71.Text
users(a, 2) = TextBox72.Text
users(a, 3) = TextBox73.Text
users(a, 4) = TextBox74.Text
users(a, 5) = ComboBox20.Value
users(a, 6) = TextBox58.Text
users(a, 7) = TextBox61.Text
users(a, 8) = TextBox64.Text
users(a, 9) = TextBox65.Text
users(a, 10) = CheckBox1.Value
users(a, 11) = CheckBox4.Value
users(a, 12) = CheckBox7.Value
users(a, 13) = ComboBox21.Value
users(a, 14) = TextBox59.Text
users(a, 15) = TextBox62.Text
users(a, 16) = TextBox66.Text
users(a, 17) = TextBox67.Text
users(a, 18) = CheckBox2.Value
users(a, 19) = CheckBox5.Value
users(a, 20) = CheckBox8.Value
users(a, 21) = ComboBox22.Value
users(a, 22) = TextBox60.Text
users(a, 23) = TextBox63.Text
users(a, 24) = TextBox68.Text
users(a, 25) = TextBox69.Text
users(a, 26) = CheckBox3.Value
users(a, 27) = CheckBox6.Value
users(a, 28) = CheckBox9.Value
users(a, 29) = ComboBox20.ListIndex
users(a, 30) = ComboBox21.ListIndex
users(a, 31) = ComboBox22.ListIndex
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Sheets("Data3").Cells.ClearContents
Sheets("Data4").Cells.ClearContents
mystr = "A1:Q99"
Sheets("Data1").Range(mystr) = itemform
mystr = "A1:E99"
Sheets("Data2").Range(mystr) = setrefill
mystr = "A1:AF999"
Sheets("Data3").Range(mystr) = users
Sheets("Data4").Range("A1").Value = TextBox9.Text
Sheets("Data4").Range("A2").Value = TextBox11.Text
Sheets("Data4").Range("A3").Value = TextBox20.Text
Sheets("Data4").Range("A4").Value = TextBox13.Text
Sheets("Data4").Range("A5").Value = TextBox21.Text
Sheets("Data4").Range("A6").Value = ComboBox17.Value
Sheets("Data4").Range("A7").Value = ComboBox18.Value
Sheets("Data4").Range("A8").Value = TextBox8.Text
Sheets("Data4").Range("A9").Value = ComboBox16.Value
Sheets("Data4").Range("A10").Value = TextBox22.Text
Sheets("Data4").Range("A11").Value = TextBox23.Text
Sheets("Data4").Range("A12").Value = TextBox24.Text
Sheets("Data4").Range("A13").Value = TextBox25.Text
Sheets("Data4").Range("A14").Value = TextBox26.Text
Sheets("Data4").Range("A15").Value = TextBox27.Text
Sheets("Data4").Range("A16").Value = TextBox28.Text
Sheets("Data4").Range("A17").Value = TextBox29.Text
Sheets("Data4").Range("A18").Value = TextBox30.Text
Sheets("Data4").Range("A19").Value = TextBox31.Text
Sheets("Data4").Range("A20").Value = TextBox32.Text
Sheets("Data4").Range("A21").Value = TextBox33.Text
Sheets("Data4").Range("A22").Value = TextBox34.Text
Sheets("Data4").Range("A23").Value = TextBox35.Text
Sheets("Data4").Range("A24").Value = TextBox36.Text
Sheets("Data4").Range("A25").Value = TextBox37.Text
Sheets("Data4").Range("A26").Value = TextBox38.Text
Sheets("Data4").Range("A27").Value = TextBox39.Text
Sheets("Data4").Range("A28").Value = TextBox40.Text
Sheets("Data4").Range("A29").Value = TextBox41.Text
Sheets("Data4").Range("A30").Value = TextBox42.Text
Sheets("Data4").Range("A31").Value = TextBox43.Text
Sheets("Data4").Range("A32").Value = TextBox44.Text
Sheets("Data4").Range("A33").Value = TextBox45.Text
Sheets("Data4").Range("A34").Value = TextBox46.Text
Sheets("Data4").Range("A35").Value = TextBox47.Text
Sheets("Data4").Range("A36").Value = TextBox48.Text
Sheets("Data4").Range("A37").Value = TextBox49.Text
Sheets("Data4").Range("A38").Value = TextBox50.Text
Sheets("Data4").Range("A39").Value = TextBox51.Text
Sheets("Data4").Range("A40").Value = TextBox52.Text
Sheets("Data4").Range("A41").Value = TextBox53.Text
Sheets("Data4").Range("A42").Value = TextBox55.Text
Sheets("Data4").Range("A43").Value = TextBox56.Text
Sheets("Data4").Range("A44").Value = TextBox57.Text
Sheets("Data4").Range("A45").Value = ComboBox19.Value
Sheets("Data4").Range("A46").Value = ComboBox26.Value
Sheets("Data4").Range("A47").Value = TextBox81.Text
Sheets("Data4").Range("A48").Value = TextBox91.Text
Sheets("Data4").Range("A49").Value = lastuser
Sheets("Data4").Range("A50").Value = lastitem
Sheets("Data4").Range("A51").Value = lastrefill
Sheets("Data4").Range("A52").Value = TextBox114.Text
Sheets("Data4").Range("A53").Value = TextBox115.Text
Sheets("Data4").Range("A54").Value = TextBox92.Text
Sheets("Data4").Range("A55").Value = TextBox93.Text
Sheets("Data4").Range("A56").Value = TextBox94.Text
Sheets("Data4").Range("A57").Value = TextBox95.Text
Sheets("Data4").Range("A58").Value = TextBox96.Text
Sheets("Data4").Range("A59").Value = TextBox97.Text
Sheets("Data4").Range("A60").Value = TextBox98.Text
Sheets("Data4").Range("A61").Value = TextBox99.Text
Sheets("Data4").Range("A62").Value = TextBox100.Text
Sheets("Data4").Range("A63").Value = TextBox101.Text
Sheets("Data4").Range("A64").Value = TextBox102.Text
Sheets("Data4").Range("A65").Value = TextBox103.Text
Sheets("Data4").Range("A66").Value = TextBox104.Text
Sheets("Data4").Range("A67").Value = TextBox105.Text
Sheets("Data4").Range("A68").Value = TextBox106.Text
Sheets("Data4").Range("A69").Value = TextBox107.Text
Sheets("Data4").Range("A70").Value = TextBox108.Text

Sheets("Data4").Range("A71").Value = TextBox117.Text
Sheets("Data4").Range("A72").Value = TextBox112.Text
Sheets("Data4").Range("A73").Value = ComboBox33.Value
Sheets("Data4").Range("A74").Value = TextBox118.Text
Sheets("Data4").Range("A75").Value = TextBox119.Text
End Sub
