VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Savings Quote Maker"
   ClientHeight    =   8985.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11610
   OleObjectBlob   =   "Savings Quote Maker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox11_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox12_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox13_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox14_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox2_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox3_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox4_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox5_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox6_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox7_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CheckBox8_Click()
If loading <> 1 And scrolling <> 1 Then Call save: Call calculate
End Sub
Private Sub CommandButton13_Click()
Call save
On Error GoTo errorhandler2
Dim oEmbFile As Object
Application.DisplayAlerts = False
Set oEmbFile = ThisWorkbook.Sheets("SA").OLEObjects(1)
    oEmbFile.Verb Verb:=xlPrimary
Set oEmbFile = Nothing
Application.DisplayAlerts = True
Dim wrddoc As Document
Application.Wait (Now + TimeValue("00:00:06"))
Set objWord = GetObject(, "Word.Application")
objWord.Visible = True
Set wrddoc = objWord.Documents("Document in " & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)))
wrddoc.Tables(1).Rows(1).Cells(2).Range.Text = "G&K SERVICES CANADA INC." & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 3, False) & " " & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 4, False) & vbNewLine & WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 5, False)
wrddoc.Shapes("Text Box 23").TextFrame.TextRange = UserForm1.CustName.Text
wrddoc.Shapes("Text Box 49").TextFrame.TextRange = UserForm1.CustNum.Text
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
a = 1
b = 3
Do While a < 199 And b < 13
If arrays(11, a) = True And arrays(10, a) <> "OS" And arrays(10, a) <> "AR" Then
wrddoc.Tables(2).Rows(b).Cells(1).Range.Text = arrays(1, a) & " " & arrays(2, a)
If arrays(9, a) = "W" Then wrddoc.Tables(2).Rows(b).Cells(5).Range.Text = "Weekly"
If arrays(9, a) = "B" Then wrddoc.Tables(2).Rows(b).Cells(5).Range.Text = "Bi-Weekly"
If arrays(9, a) = "M" Then wrddoc.Tables(2).Rows(b).Cells(5).Range.Text = "Monthly"
wrddoc.Tables(2).Rows(b).Cells(4).Range.Text = arrays(6, a)
wrddoc.Tables(2).Rows(b).Cells(6).Range.Text = arrays(10, a)
wrddoc.Tables(2).Rows(b).Cells(2).Range.Text = arrays(3, a)
wrddoc.Tables(2).Rows(b).Cells(3).Range.Text = arrays(0, a)
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
If err2 = 0 Then
err2 = 1
Set wrddoc = objWord.Documents("Document in " & ThisWorkbook.Name)
Resume Next
End If
MsgBox "13: An error has occured linking with the Word document. Please close Word and try again."
Exit Sub
End Sub
Private Sub CommandButton14_Click()
Call save
On Error GoTo errorhandler2
Dim oEmbFile As Object
Application.DisplayAlerts = False
Set oEmbFile = ThisWorkbook.Sheets("SA").OLEObjects(2)
    oEmbFile.Verb Verb:=xlPrimary
Set oEmbFile = Nothing
Application.DisplayAlerts = True
Dim wrddoc As Document
Application.Wait (Now + TimeValue("00:00:06"))
Set objWord = GetObject(, "Word.Application")
objWord.Visible = True
Set wrddoc = objWord.Documents("Document in " & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)))
wrddoc.Shapes("Text Box 5").TextFrame.TextRange = UserForm1.CustName.Text
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
Do While a < 199 And b < 32
If arrays(11, a) = True And arrays(10, a) <> "OS" And arrays(10, a) <> "AR" Then
If b > 12 Then
wrddoc.Tables(2).Rows(b - 10).Cells(1).Range.Text = arrays(1, a) & " " & arrays(2, a)
If arrays(9, a) = "W" Then wrddoc.Tables(2).Rows(b - 10).Cells(5).Range.Text = "Weekly"
If arrays(9, a) = "B" Then wrddoc.Tables(2).Rows(b - 10).Cells(5).Range.Text = "Bi-Weekly"
If arrays(9, a) = "M" Then wrddoc.Tables(2).Rows(b - 10).Cells(5).Range.Text = "Monthly"
wrddoc.Tables(2).Rows(b - 10).Cells(4).Range.Text = arrays(6, a)
wrddoc.Tables(2).Rows(b - 10).Cells(6).Range.Text = arrays(10, a)
wrddoc.Tables(2).Rows(b - 10).Cells(2).Range.Text = arrays(3, a)
wrddoc.Tables(2).Rows(b - 10).Cells(3).Range.Text = arrays(0, a)
End If
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
errorhandler2:
If err2 = 0 Then
err2 = 1
Set wrddoc = objWord.Documents("Document in " & ThisWorkbook.Name)
Resume Next
End If
MsgBox "13: An error has occured linking with the Word document. Please close Word and try again."
Exit Sub
End Sub

Private Sub CommandButton15_Click()
If CommandButton15.Caption = 4 Then
MultiPage1.Width = 710
Label154.Visible = True
Label164.Visible = True
UserForm1.Width = 725
CommandButton15.Caption = 3
For a = 1 To 8
Controls("Item" & a & 2).Visible = True
Controls("Item" & a & 10).Visible = True
Next
Else
Label154.Visible = False
Label164.Visible = False
MultiPage1.Width = 570
UserForm1.Width = 585
CommandButton15.Caption = 4
For a = 1 To 8
Controls("Item" & a & 2).Visible = False
Controls("Item" & a & 10).Visible = False
Next
End If
End Sub
Private Sub CommandButton16_Click()
Unload Me
Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
Erase arrays
UserForm1.Show
End Sub

Private Sub CreateButton_Click()
Call save
newrw = 13
summaryline = 43
ThisWorkbook.Sheets("Sheet4").Visible = True
ThisWorkbook.Sheets("Sheet4").Copy
ThisWorkbook.Sheets("Sheet4").Visible = False
Dim ws As Worksheet
Set ws = ActiveWorkbook.Worksheets("Sheet4")
ws.Cells(3, 6).Value = [TODAY()]
ws.Cells(3, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 3, False)
ws.Cells(4, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 4, False)
ws.Cells(5, 1).Value = WorksheetFunction.VLookup(Val(Left(ComboBox33.Value, 3)), ThisWorkbook.Worksheets("DPCs").Range("A:E"), 5, False)
ws.Cells(48, 1).Value = "Prices are based on a " & TextBox103.Text & " month term."
ws.Cells(7, 1).Value = "Prepared by: " & TextBox112.Text
ws.Cells(9, 1).Value = CustName.Text
ws.Cells(9, 7).Value = CustNum.Text
aa = 1
weekly = 0
monthly = 0
biweekly = 0
Dim weekly2(198) As Integer
Dim biweekly2(198) As Integer
Dim monthly2(198) As Integer
Do While aa < 199
If arrays(11, aa) = True And arrays(0, aa) > 0 Then
If arrays(9, aa) = "W" Then weekly = weekly + 1: weekly2(weekly) = aa
If arrays(9, aa) = "B" Then biweekly = biweekly + 1: biweekly2(biweekly) = aa
If arrays(9, aa) = "M" Then monthly = monthly + 1: monthly2(monthly) = aa
End If
aa = aa + 1
Loop
weeklys = 0
biweeklys = 0
monthlys = 0
bb = 1
cc = 1
dd = 1
If weekly > 0 Then ws.Cells(newrw, 1) = "WEEKLY:": ws.Range("A" & newrw).Font.Bold = True: newrw = newrw + 1
back1:
Do While bb <= weekly
aa = weekly2(bb)
weeklys = weeklys + (Round(arrays(4, aa) * arrays(3, aa), 2) - Round(arrays(6, aa) * arrays(3, aa), 2))
bb = bb + 1
GoTo insert
Loop
If biweekly > 0 And cc = 1 Then ws.Cells(newrw, 1) = "BIWEEKLY:": ws.Range("A" & newrw).Font.Bold = True: newrw = newrw + 1
Do While cc <= biweekly
aa = biweekly2(cc)
biweeklys = biweeklys + ((Round(arrays(4, aa) * arrays(3, aa), 2) - Round(arrays(6, aa) * arrays(3, aa), 2)) / 2)
cc = cc + 1
GoTo insert
Loop
If monthly > 0 And dd = 1 Then ws.Cells(newrw, 1) = "MONTHLY:": ws.Range("A" & newrw).Font.Bold = True: newrw = newrw + 1
Do While dd <= monthly
aa = monthly2(dd)
monthlys = monthlys + ((Round(arrays(4, aa) * arrays(3, aa), 2) - Round(arrays(6, aa) * arrays(3, aa), 2)) / 4)
dd = dd + 1
GoTo insert
Loop
GoTo done
insert:
If newrw > 32 Then ws.Rows(newrw).EntireRow.insert: ws.Rows(newrw).EntireRow.insert: summaryline = summaryline + 2
ws.Cells(newrw, 1) = arrays(1, aa)
ws.Range("A" & newrw).Font.Bold = False
ws.Range("A" & newrw + 1).Font.Bold = False
ws.Cells(newrw, 2) = arrays(3, aa)
ws.Cells(newrw, 3) = arrays(4, aa)
ws.Cells(newrw, 4) = "=B" & newrw & "*C" & newrw
ws.Cells(newrw, 5) = arrays(6, aa)
ws.Cells(newrw, 6) = "=B" & newrw & "*E" & newrw
ws.Cells(newrw, 7) = "=D" & newrw & "-F" & newrw
ws.Cells(newrw + 1, 1) = "         " & arrays(2, aa)
newrw = newrw + 2
GoTo back1
done:
newrw = newrw + 1
Dim savings As Double
If CheckBox11.Value = True Then ws.Cells(newrw, 1) = "Energy Surcharge": ws.Cells(newrw, 4) = Energy2.Text: ws.Cells(newrw, 6) = Energy4.Text: ws.Cells(newrw, 7) = "=D" & newrw & "-F" & newrw: newrw = newrw + 1: savings = Energy5.Text
If CheckBox12.Value = True Then ws.Cells(newrw, 1) = "Enviro Surcharge": ws.Cells(newrw, 4) = Enviro2.Text: ws.Cells(newrw, 6) = Enviro4.Text: ws.Cells(newrw, 7) = "=D" & newrw & "-F" & newrw: newrw = newrw + 1: savings = savings + Enviro5.Text
If CheckBox13.Value = True Then ws.Cells(newrw, 1) = "Image Guard": ws.Cells(newrw, 4) = Image1.Text: ws.Cells(newrw, 6) = Image3.Text: ws.Cells(newrw, 7) = "=D" & newrw & "-F" & newrw: newrw = newrw + 1: savings = savings + Image4.Text
If CheckBox14.Value = True Then ws.Cells(newrw, 1) = "Prep Guard": ws.Cells(newrw, 4) = Prep1.Text: ws.Cells(newrw, 6) = Prep3.Text: ws.Cells(newrw, 7) = "=D" & newrw & "-F" & newrw: newrw = newrw + 1: savings = savings + Prep4.Text
If weekly > 0 Then
savings = savings + weeklys + monthlys + biweeklys
ElseIf biweekly > 0 Then
savings = (savings / 2) + monthlys + biweeklys
Else
savings = (savings / 4) + monthlys
End If
If savings = 0 Then ws.Cells(summaryline, 2) = "=G" & summaryline + 1 Else ws.Cells(summaryline, 2) = Round(savings * 1.13, 2)
UserForm1.Hide
End Sub
Private Sub Image2_Change()
If loading <> 1 Then Call save: Call calculate
End Sub
Private Sub Image5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
MsgBox ("Made by Andrew Aaron - aaaron@gkservices.com - last updated February 6, 2015")
End Sub
Private Sub Importbutton_Click()
filetoOpen = Application.GetOpenFilename("Excel File (*.xls), *.xls")
If filetoOpen = False Then Exit Sub
If Len(Dir(filetoOpen)) = 0 Then Exit Sub
Dim bWasClosed As Boolean
Dim sPath As String
Dim wbSource As Workbook
Dim wbDest As Workbook
Set wbDest = ActiveWorkbook
On Error Resume Next ' avoid the error if not open
Set wbSource = filetoOpen
On Error GoTo errorsht
If wbSource Is Nothing Then
   bWasClosed = True
   Set wbSource = Application.Workbooks.Open("" & filetoOpen & "")
End If
Application.DisplayAlerts = False
Dim data As Worksheet
On Error GoTo errorsht
Set data = wbSource.Sheets("Sheet1")
On Error GoTo 0
GoTo news
errorsht:
MsgBox "Invalid Format"
Exit Sub
news:
rw = 1
Do While rw < 8
If IsNumeric(data.Cells(rw, 30)) And data.Cells(rw, 30) <> "" Then GoTo start
rw = rw + 1
Loop
MsgBox "Invalid Format"
Exit Sub
start:
Erase arrays, image, prep
loading = 1
ScrollBar1.Value = 1
lastrw = rw
freq = data.Cells(lastrw, 32)
options2 = freq
Do While data.Cells(lastrw, 30) <> ""
lastrw = lastrw + 1
Loop
data.Range("A" & rw & ":AG" & lastrw).Sort Key1:=data.Columns("AF"), Order1:=xlAscending, Key2:=data.Columns("N"), Order2:=xlAscending, Order3:=xlAscending, Key3:=data.Columns("L")
rw2 = rw
options2 = data.Cells(rw2, 32)
current = options2
Do While rw2 < lastrw
If data.Cells(rw2, 32) <> current And data.Cells(rw2, 30) <> 0 And data.Cells(rw2, 22) <> 0 Then options2 = options2 & data.Cells(rw2, 32)
current = data.Cells(rw2, 32)
rw2 = rw2 + 1
Loop
CustName.Text = data.Cells(rw, 3)
CustNum.Text = data.Cells(rw, 4) & "-" & data.Cells(rw, 5)
newrw = 1
amt = 0
Do While data.Cells(rw, 30) <> ""
nextrw:
If data.Cells(rw, 24) = "Y" Then
osci = data.Cells(rw, 30)
ci = 0
Else
ci = data.Cells(rw, 30)
osci = 0
End If
osibilled = 0
ibilled = 0
amt = 0
osamt = 0
here:
isc = data.Cells(rw, 14) & "-" & data.Cells(rw, 16) & "-" & data.Cells(rw, 18)
st = data.Cells(rw, 14) & "-" & data.Cells(rw, 16)
bm = data.Cells(rw, 21)
up = data.Cells(rw, 22)
wn = data.Cells(rw, 6)
ir = data.Cells(rw, 12)
desc = data.Cells(rw, 15)
osup = data.Cells(rw, 25)
arp = data.Cells(rw, 31)
repp = data.Cells(rw, 23)
freq = data.Cells(rw, 32)
billingpct = data.Cells(rw, 33)
If data.Cells(rw, 24) = "Y" Then
If data.Cells(rw, 30) Mod 2 = 0 Then
osamt = osamt + ((data.Cells(rw, 30) / 2) * osup)
osibilled = osibilled + (data.Cells(rw, 30) / 2)
Else
osamt = osamt + (((data.Cells(rw, 30) - 1) / 2) * osup)
osibilled = osibilled + ((data.Cells(rw, 30) - 1) / 2)
End If
Else
If data.Cells(rw, 30) Mod 2 = 0 Then
amt = amt + ((data.Cells(rw, 30) / 2) * up)
ibilled = ibilled + (data.Cells(rw, 30) / 2)
Else
amt = amt + (((data.Cells(rw, 30) - 1) / 2) * up)
ibilled = ibilled + ((data.Cells(rw, 30) - 1) / 2)
End If
End If
rw = rw + 1
If data.Cells(rw, 12) = ir Then
If data.Cells(rw, 24) = "Y" Then osci = osci + data.Cells(rw, 30) Else: ci = ci + data.Cells(rw, 30)
GoTo here
ElseIf (ci > 0 Or osci > 0) And up > 0 Then
arrays(9, newrw) = freq
'If ci = 0 Then
'ci = osci
'amt = osamt
'up = osup
'ibilled = osibilled
'osflag = " OUTSIZES"
'End If
DOOVER:
try = ""
On Error Resume Next
try = WorksheetFunction.VLookup(st, ThisWorkbook.Worksheets("Lists").Range("A:B"), 2, False)
If try = "" Then
arrays(1, newrw) = desc & osflag
Else: arrays(1, newrw) = try & osflag
End If
arrays(11, newrw) = True
arrays(2, newrw) = isc
If wn <> "" And bm = "FR" And asked <> 1 Then
    If MsgBox("Convert garments to CI pricing?", vbYesNo) = vbNo Then frbilling = 1
    asked = 1
End If

If bm = "US" Then
arrays(3, newrw) = Round(ci * (billingpct / 100), 0)
ElseIf Left(isc, 4) < 2099 And Left(isc, 4) > 2000 Then
arrays(3, newrw) = ci / 2
Else: arrays(3, newrw) = ci
End If
If wn <> "" And bm = "FR" Then
If frbilling = 1 Then
arrays(3, newrw) = ibilled
Else: arrays(3, newrw) = ci: If ci = 0 Then up = Round(osamt / osci / 1.35, 2) Else up = Round(amt / ci, 2)
'does this work
End If
End If
arrays(4, newrw) = up
arrays(6, newrw) = up
arrays(0, newrw) = ci
If osflag = "" Then arrays(10, newrw) = repp Else arrays(10, newrw) = "OS"
arrays(9, newrw) = freq
'arrays
'0CI
'1description
'2isc
'3Qty
'4current price
'5current total
'6new price
'7new total
'8savings
'9frequency
'10replacement
'11check box
newrw = newrw + 1
End If
If osflag = "" And osci > 0 Then
ci = osci
amt = osamt
up = osup
ibilled = osibilled
osflag = " OUTSIZES"
GoTo DOOVER
Else: osflag = ""
End If
If arp > 0.99 Then
try = ""
On Error Resume Next
try = WorksheetFunction.VLookup(st, Worksheets("Lists").Range("A:B"), 2, False)
If try = "" Then
arrays(1, newrw) = desc & " AUTO-REPLACE " & arp & "%"
Else: arrays(1, newrw) = try & " AUTO-REPLACE " & arp & "%"
End If
arrays(11, newrw) = True
arrays(2, newrw) = isc
arrays(3, newrw) = Round(ci * (arp / 100), 0)
arrays(4, newrw) = repp
arrays(6, newrw) = repp
arrays(10, newrw) = "AR"
arrays(9, newrw) = freq
newrw = newrw + 1
End If
Loop
If bWasClosed Then wbSource.Close False  ' close without saving
aa = 1
For aa = 1 To 8
For bb = 0 To 10
If bb = 9 And arrays(bb, aa) = "" Then Controls("Item" & aa & bb).Text = "W" Else Controls("Item" & aa & bb).Text = arrays(bb, aa)
Next
If arrays(11, aa) = True Then Controls("Checkbox" & aa).Value = True Else Controls("Checkbox" & aa).Value = False
If SafeVlookup(Left(arrays(2, aa), 8), ThisWorkbook.Worksheets("Sheet2").Range("H:I"), 2, False, 0) > 0 Or arrays(10, aa) = "OS" Then Controls("Info" & aa).Visible = True Else Controls("info" & aa).Visible = False
Next
Call calculate
loading = 0
End Sub
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
Public Sub calculate()
       If loading2 <> 1 Then
    loading2 = 1
    aa = 1
For aa = 1 To 8
For bb = 0 To 10
arrays(bb, aa + UserForm1.CheckBox1.Caption - 1) = UserForm1.Controls("Item" & aa & bb).Text
Next
arrays(11, aa + UserForm1.CheckBox1.Caption - 1) = UserForm1.Controls("Checkbox" & aa).Value
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
Sub iclick(ByVal x As Double)
If Controls("Item" & x & 10).Text = "OS" Then
load2 = 1
UserForm3.Label3.Caption = Controls("Item" & x & 1).Text
If x = 1 Then UserForm3.TextBox1 = arrays(6, ScrollBar1.Value - 1) Else UserForm3.TextBox1 = Controls("Item" & x - 1 & 6).Text
UserForm3.TextBox3 = Controls("Item" & x & 6).Text
UserForm3.TextBox2.Text = Round(((Val(UserForm3.TextBox3.Text) / Val(UserForm3.TextBox1.Text)) - 1) * 100, 1)
iform = x
load2 = 0
UserForm3.Show
Else
load2 = 1
UserForm2.Label3.Caption = Controls("Item" & x & 1).Text
sqft = SafeVlookup(Left(Controls("Item" & x & 2).Text, 8), ThisWorkbook.Worksheets("Sheet2").Range("H:I"), 2, False, 0)
UserForm2.TextBox1 = Round(Controls("Item" & x & 6).Text / sqft, 3)
UserForm2.TextBox2 = Controls("Item" & x & 6).Text
iform = x
load2 = 0
UserForm2.Show
End If
End Sub
Private Sub Info1_Click()
Call iclick(1)
End Sub
Private Sub Info2_Click()
Call iclick(2)
End Sub
Private Sub Info3_Click()
Call iclick(3)
End Sub
Private Sub Info4_Click()
Call iclick(4)
End Sub
Private Sub Info5_Click()
Call iclick(5)
End Sub
Private Sub Info6_Click()
Call iclick(6)
End Sub
Private Sub Info7_Click()
Call iclick(7)
End Sub
Private Sub Info8_Click()
Call iclick(8)
End Sub
Private Sub MultiPage1_Change()
If MultiPage1.Value = 1 Then
If OptionButton3.Value = True Then TextBox106.Text = Energy3.Text & "%" Else TextBox106.Text = "$" & Energy3.Text
If OptionButton7.Value = True Then TextBox107.Text = Enviro3.Text & "%" Else TextBox107.Text = "$" & Enviro3.Text
TextBox111.Text = IIf(CheckBox13.Value = True, "YES", "NO") & "/" & IIf(CheckBox14.Value = True, "YES", "NO")
End If
End Sub
Private Sub OptionButton1_Change()
Call calculate
End Sub
Private Sub OptionButton3_Change()
Call calculate
End Sub
Private Sub OptionButton5_Change()
Call calculate
End Sub
Private Sub OptionButton7_Change()
Call calculate
End Sub
Private Sub Prep2_Change()
If loading <> 1 Then Call save: Call calculate
End Sub
Private Sub ScrollBar1_Change()
If loading <> 1 Then
scrolling = 1
Call save
For i = 1 To 8
Controls("Checkbox" & i).Caption = ScrollBar1.Value + i - 1
Next
aa = 1
For aa = 1 To 8
For bb = 0 To 10
If bb = 9 And arrays(bb, aa + ScrollBar1.Value - 1) = "" Then Controls("Item" & aa & bb).Text = "W" Else Controls("Item" & aa & bb).Text = arrays(bb, aa + ScrollBar1.Value - 1)
Next
If arrays(11, aa + ScrollBar1.Value - 1) = True Then Controls("Checkbox" & aa).Value = True Else Controls("Checkbox" & aa).Value = False
If SafeVlookup(Left(arrays(2, aa + ScrollBar1.Value - 1), 8), ThisWorkbook.Worksheets("Sheet2").Range("H:I"), 2, False, 0) > 0 Or (arrays(10, aa + ScrollBar1.Value - 1) = "OS" And arrays(2, aa + ScrollBar1.Value - 1) = arrays(2, aa + ScrollBar1.Value - 2)) Then Controls("Info" & aa).Visible = True Else Controls("Info" & aa).Visible = False
Next
scrolling = 0
Call calculate
End If
End Sub
Sub save()
If loading <> 1 Then
aa = 1
For aa = 1 To 8
For bb = 0 To 10
arrays(bb, aa + CheckBox1.Caption - 1) = Controls("Item" & aa & bb).Text
Next
arrays(11, aa + CheckBox1.Caption - 1) = Controls("Checkbox" & aa).Value
Next
End If
End Sub
Private Sub Updatebutton_Click()
loading = 1
aa = 1
For aa = 1 To 198
If arrays(6, aa) <> "" Then arrays(6, aa) = Round(arrays(4, aa) - (arrays(4, aa) * (Val(Reduce.Text) / 100)), 2)
Next
aa = 1
For aa = 1 To 8
Controls("Item" & aa & 6).Text = arrays(6, aa + ScrollBar1.Value - 1)
Next
loading = 0
Call calculate
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    UserForm1.Hide
  End If
  Sheets("Data1").Cells.ClearContents
Sheets("Data2").Cells.ClearContents
mystr = "A1:GP12"
Sheets("Data1").Range(mystr) = arrays

Sheets("Data2").Range("A1").Value = CustName.Text
Sheets("Data2").Range("A2").Value = CustNum.Text
Sheets("Data2").Range("A3").Value = Energy1.Text
Sheets("Data2").Range("A4").Value = Energy3.Text
Sheets("Data2").Range("A5").Value = Enviro1.Text
Sheets("Data2").Range("A6").Value = Enviro3.Text
Sheets("Data2").Range("A7").Value = OptionButton1.Value
Sheets("Data2").Range("A8").Value = OptionButton3.Value
Sheets("Data2").Range("A9").Value = OptionButton5.Value
Sheets("Data2").Range("A10").Value = OptionButton7.Value
Sheets("Data2").Range("A11").Value = Image1.Text
Sheets("Data2").Range("A12").Value = Prep1.Text
Sheets("Data2").Range("A13").Value = Image2.Value
Sheets("Data2").Range("A14").Value = Prep2.Value
Sheets("Data2").Range("A15").Value = ComboBox33.Value
Sheets("Data2").Range("A16").Value = TextBox103.Text
Sheets("Data2").Range("A17").Value = TextBox112.Text
Sheets("Data2").Range("A18").Value = CheckBox11.Value
Sheets("Data2").Range("A19").Value = CheckBox12.Value
Sheets("Data2").Range("A20").Value = CheckBox13.Value
Sheets("Data2").Range("A21").Value = CheckBox14.Value
Sheets("Data2").Range("A22").Value = TextBox92.Text
Sheets("Data2").Range("A23").Value = TextBox109.Text
Sheets("Data2").Range("A24").Value = TextBox110.Text
Sheets("Data2").Range("A25").Value = TextBox93.Text
Sheets("Data2").Range("A26").Value = TextBox94.Text
Sheets("Data2").Range("A27").Value = TextBox95.Text
Sheets("Data2").Range("A28").Value = TextBox96.Text
Sheets("Data2").Range("A29").Value = TextBox97.Text
Sheets("Data2").Range("A30").Value = TextBox111.Text
Sheets("Data2").Range("A31").Value = TextBox106.Text
Sheets("Data2").Range("A32").Value = TextBox107.Text
Sheets("Data2").Range("A33").Value = TextBox98.Text
Sheets("Data2").Range("A34").Value = TextBox99.Text
Sheets("Data2").Range("A35").Value = TextBox100.Text
Sheets("Data2").Range("A36").Value = TextBox101.Text
Sheets("Data2").Range("A37").Value = TextBox104.Text
Sheets("Data2").Range("A38").Value = TextBox105.Text
Sheets("Data2").Range("A39").Value = TextBox108.Text
Sheets("Data2").Range("A40").Value = TextBox102.Text
Sheets("Data2").Range("A41").Value = OptionButton2.Value
Sheets("Data2").Range("A42").Value = OptionButton4.Value
Sheets("Data2").Range("A43").Value = OptionButton6.Value
Sheets("Data2").Range("A44").Value = OptionButton8.Value

End Sub
Private Sub UserForm_Initialize()
a = 1
loading = 1
For a = 1 To 8
With Controls("Item" & a & "9")
.AddItem "W"
.AddItem "B"
.AddItem "M"
.ListIndex = 0
End With
Next
With Image2
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
.AddItem "TB"
.AddItem "TB2"
.AddItem "TS1"
.AddItem "TS"
.AddItem "TS2"
.AddItem "TG1"
.AddItem "TG"
.AddItem "TG2"
.AddItem "TP1"
.AddItem "TP"
.AddItem "TP2"
.AddItem "TP3"
.AddItem "TP4"
.ListIndex = 5
End With
With Prep2
.AddItem "Chrome"
.AddItem "Copper"
.AddItem "Bronze"
.AddItem "Jade"
.AddItem "Silver"
.AddItem "Gold"
.AddItem "Platinum"
.AddItem "Diamond"
.ListIndex = 2
End With
dp = 1
DPCs = Worksheets("DPCs").UsedRange.Rows.Count
Do While dp <= DPCs
ComboBox33.AddItem (Sheets("DPCs").Range("A" & dp).Value & "-" & Sheets("DPCs").Range("B" & dp).Value)
dp = dp + 1
Loop
ComboBox33.ListIndex = 1

For p = 0 To 198
For j = 0 To 11
arrays(j, p) = ThisWorkbook.Sheets("Data1").Cells(j + 1, p + 1).Value
Next
Next

aa = 1
For aa = 1 To 8
For bb = 0 To 10
If bb = 9 And arrays(bb, aa) = "" Then Controls("Item" & aa & bb).Text = "W" Else Controls("Item" & aa & bb).Text = arrays(bb, aa)
Next
If arrays(11, aa) = True Then Controls("Checkbox" & aa).Value = True Else Controls("Checkbox" & aa).Value = False
If SafeVlookup(Left(arrays(2, aa), 8), ThisWorkbook.Worksheets("Sheet2").Range("H:I"), 2, False, 0) > 0 Or (arrays(10, aa) = "OS" And arrays(2, aa) = arrays(2, aa - 1)) Then Controls("Info" & aa).Visible = True Else Controls("Info" & aa).Visible = False
Next
If Sheets("Data2").Range("A21").Value <> "" Then
CustName.Text = Sheets("Data2").Range("A1").Value
CustNum.Text = Sheets("Data2").Range("A2").Value
Energy1.Text = Sheets("Data2").Range("A3").Value
Energy3.Text = Sheets("Data2").Range("A4").Value
Enviro1.Text = Sheets("Data2").Range("A5").Value
Enviro3.Text = Sheets("Data2").Range("A6").Value
OptionButton1.Value = Sheets("Data2").Range("A7").Value
OptionButton3.Value = Sheets("Data2").Range("A8").Value
OptionButton5.Value = Sheets("Data2").Range("A9").Value
OptionButton7.Value = Sheets("Data2").Range("A10").Value
Image1.Text = Sheets("Data2").Range("A11").Value
Prep1.Text = Sheets("Data2").Range("A12").Value
Image2.Value = Sheets("Data2").Range("A13").Value
Prep2.Value = Sheets("Data2").Range("A14").Value
ComboBox33.Value = Sheets("Data2").Range("A15").Value
TextBox103.Text = Sheets("Data2").Range("A16").Value
TextBox112.Text = Sheets("Data2").Range("A17").Value
CheckBox11.Value = Sheets("Data2").Range("A18").Value
CheckBox12.Value = Sheets("Data2").Range("A19").Value
CheckBox13.Value = Sheets("Data2").Range("A20").Value
CheckBox14.Value = Sheets("Data2").Range("A21").Value
TextBox92.Text = Sheets("Data2").Range("A22").Value
TextBox109.Text = Sheets("Data2").Range("A23").Value
TextBox110.Text = Sheets("Data2").Range("A24").Value
TextBox93.Text = Sheets("Data2").Range("A25").Value
TextBox94.Text = Sheets("Data2").Range("A26").Value
TextBox95.Text = Sheets("Data2").Range("A27").Value
TextBox96.Text = Sheets("Data2").Range("A28").Value
TextBox97.Text = Sheets("Data2").Range("A29").Value
TextBox111.Text = Sheets("Data2").Range("A30").Value
TextBox106.Text = Sheets("Data2").Range("A31").Value
TextBox107.Text = Sheets("Data2").Range("A32").Value
TextBox98.Text = Sheets("Data2").Range("A33").Value
TextBox99.Text = Sheets("Data2").Range("A34").Value
TextBox100.Text = Sheets("Data2").Range("A35").Value
TextBox101.Text = Sheets("Data2").Range("A36").Value
TextBox104.Text = Sheets("Data2").Range("A37").Value
TextBox105.Text = Sheets("Data2").Range("A38").Value
TextBox108.Text = Sheets("Data2").Range("A39").Value
TextBox102.Text = Sheets("Data2").Range("A40").Value
OptionButton2.Value = Sheets("Data2").Range("A41").Value
OptionButton4.Value = Sheets("Data2").Range("A42").Value
OptionButton6.Value = Sheets("Data2").Range("A43").Value
OptionButton8.Value = Sheets("Data2").Range("A44").Value
End If
Call calculate
loading = 0
 Dim TBCount As Integer
    Dim Ctrl As Control
    TBCount = 0
    For Each Ctrl In UserForm1.Controls
        If TypeName(Ctrl) = "TextBox" Then
            TBCount = TBCount + 1
            ReDim Preserve TBs(1 To TBCount)
            Set TBs(TBCount).TBGroup = Ctrl
        End If
    Next Ctrl
End Sub
