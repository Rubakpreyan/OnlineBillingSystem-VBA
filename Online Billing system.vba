Option Explicit

Dim firstnum As Double
Dim secondnum As Double
Dim answer As Double
Dim operator As String

Const apple = 12
Const orange = 9.5
Const rice = 8.5
Const tax = 17.5
Const delivery = 9.99
Const mileage = 0.55
Dim cCost As Double




Private Sub CheckBox1_Click()

End Sub

Private Sub cmdAddItem_Click()
Dim wks As Worksheet
Dim AddNew As Range

Set wks = Sheet1
Set AddNew = wks.Range("A65356").End(xlUp).Offset(1, 0)

AddNew.Offset(0, 0).Value = txtSalesNo.Value
AddNew.Offset(0, 1).Value = txtName.Value
AddNew.Offset(0, 2).Value = txtAddress.Value
AddNew.Offset(0, 3).Value = txtPostCode.Value
AddNew.Offset(0, 4).Value = txtPhone.Value
AddNew.Offset(0, 5).Value = txtApple.Value
AddNew.Offset(0, 6).Value = txtOrange.Value
AddNew.Offset(0, 7).Value = txtRice.Value
AddNew.Offset(0, 8).Value = txtDelivery.Value
AddNew.Offset(0, 9).Value = txtMileage.Value
AddNew.Offset(0, 10).Value = lblSubTotal.Value
AddNew.Offset(0, 11).Value = lblTax.Value
AddNew.Offset(0, 12).Value = lblTotal.Value

txtApple.Text = ""
txtSalesNo.Text = ""
txtName.Text = ""
txtPostCode.Text = ""
txtSalesNo.Text = ""
txtPhone.Text = ""
txtAddress.Text = ""
txtRice.Text = ""
txtOrange.Text = ""
txtDelivery.Text = ""
txtMileage.Text = ""
lblSubTotal.Value = ""
lblTax.Value = ""
lblTotal.Value = ""
txtCostOfMileage.Text = ""
txtCostOfDelivery.Text = ""
txtCostOfItems.Text = ""


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
txtDisplay.Text = "0"
lblTotal.Value = "0"
chkApple.Value = False
txtApple.Text = ""
lblSubTotal.Value = "0"
chkTax.Value = False
chkDelivery.Value = False
lblTax.Value = "0"
txtSalesNo.Text = ""
txtName.Text = ""
txtPostCode.Text = ""
txtSalesNo.Text = ""
txtPhone.Text = ""
txtAddress.Text = ""
txtCostOfMileage.Text = ""
txtCostOfDelivery.Text = ""
txtCostOfItems.Text = ""
chkRice.Value = False
txtRice.Text = ""
chkOrange.Value = False
txtOrange.Text = ""
txtDelivery.Text = ""
txtMileage.Text = ""



End Sub

Private Sub cmdTotal_Click()

If chkRice.Value = True And chkApple.Value = True And chkOrange.Value = True And chkTax.Value = True Then
cCost = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange) + (Val(txtRice.Text) * rice) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
lblSubTotal.Value = cCost
lblTax.Value = (cCost * tax / 100)
txtCostOfItems.Text = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange) + (Val(txtRice.Text) * rice)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text


ElseIf chkApple.Value = True And chkOrange.Value = True And chkTax.Value = True Then
cCost = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
'cCost = (val(txtApple.Text)*apple) + (val(txtOrange.Text)*orange)+(cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = (cCost * tax / 100)
txtCostOfItems.Text = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text


ElseIf chkApple.Value = True And chkOrange.Value = True Then
cCost = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)

'cCost = (val(txtApple.Text)*apple) + (val(txtOrange.Text)*orange)+(cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = ""
txtCostOfItems.Text = (Val(txtApple.Text) * apple) + (Val(txtOrange.Text) * orange)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text


ElseIf chkApple.Value = True And chkTax.Value = True Then
cCost = (Val(txtApple.Text) * apple) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
cCost = (Val(txtApple.Text) * apple) + (cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = (cCost * tax / 100)
txtCostOfItems.Text = (Val(txtApple.Text) * apple)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text

ElseIf chkApple.Value = True And chkTax.Value = False Then
cCost = (Val(txtApple.Text) * apple) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
'cCost = (val(txtApple.Text)*apple) + (val(txtOrange.Text)*orange)+(cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = ""
txtCostOfItems.Text = (Val(txtApple.Text) * apple)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text

ElseIf chkOrange.Value = True And chkTax.Value = True Then
cCost = (Val(txtOrange.Text) * orange) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
cCost = (Val(txtOrange.Text) * orange) + (cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = (cCost * tax / 100)
txtCostOfItems.Text = (Val(txtOrange.Text) * orange)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text

ElseIf chkOrange.Value = True And chkTax.Value = False Then
cCost = (Val(txtOrange.Text) * orange) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
'cCost = (val(txtApple.Text)*apple) + (val(txtOrange.Text)*orange)+(cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = ""
txtCostOfItems.Text = (Val(txtOrange.Text) * orange)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text


ElseIf chkRice.Value = True And chkTax.Value = True Then
cCost = (Val(txtRice.Text) * rice) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
cCost = (Val(txtRice.Text) * rice) + (cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = (cCost * tax / 100)
txtCostOfItems.Text = (Val(txtRice.Text) * rice)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text

ElseIf chkRice.Value = True And chkTax.Value = False Then
cCost = (Val(txtRice.Text) * rice) + (Val(txtDelivery.Text) * delivery) + (Val(txtMileage.Text) * mileage)
cCost = (Val(txtRice.Text) * rice) + (cCost * tax / 100)
lblSubTotal.Value = cCost
lblTax.Value = ""
txtCostOfItems.Text = (Val(txtRice.Text) * rice)
txtCostOfDelivery.Text = Val(txtDelivery.Text) * delivery
txtCostOfMileage.Text = Val(txtMileage.Text) * mileage
lblTotal.Value = Val(lblTax.Value) + Val(lblSubTotal.Value)
txtCostOfItems.Text = Format(txtCostOfItems.Text, "$#,##0.00")
lblSubTotal.Value = Format(lblSubTotal.Value, "$#,##0.00")
lblTax.Value = Format(lblTax.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
lblTotal.Value = Format(lblTotal.Value, "$#,##0.00")
txtCostOfDelivery.Text = Format(txtCostOfDelivery.Text, "$#,##0.00")
txtCostOfMileage.Text = Format(txtCostOfMileage.Text, "$#,##0.00")
txtSalesNo.Text = Evaluate("RANDBETWEEN(001,99999999)")
txtSalesNo.Text = txtSalesNo.Text + "_" + txtPostCode.Text

End If

cmdAddItem.Enabled = True







End Sub

Private Sub cmdView_Click()
Unload Me
Sheet1.Select
Sheet1.Range("A1").Select

End Sub

Private Sub CommandButton1_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "0"
Else
txtDisplay.Text = txtDisplay.Text + "0"
End If
End Sub

Private Sub CommandButton10_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "5"
Else
txtDisplay.Text = txtDisplay.Text + "5"
End If

End Sub

Private Sub CommandButton11_Click()
firstnum = txtDisplay.Text
txtDisplay.Text = ""
operator = "*"
End Sub

Private Sub CommandButton12_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "6"
Else
txtDisplay.Text = txtDisplay.Text + "6"
End If

End Sub

Private Sub CommandButton13_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "7"
Else
txtDisplay.Text = txtDisplay.Text + "7"
End If

End Sub

Private Sub CommandButton14_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "8"
Else
txtDisplay.Text = txtDisplay.Text + "8"
End If
End Sub

Private Sub CommandButton15_Click()
firstnum = txtDisplay.Text
txtDisplay.Text = ""
operator = "+"

End Sub

Private Sub CommandButton16_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "9"
Else
txtDisplay.Text = txtDisplay.Text + "9"
End If

End Sub

Private Sub CommandButton17_Click()
firstnum = txtDisplay.Text
txtDisplay.Text = ""
operator = "%"
End Sub

Private Sub CommandButton18_Click()
txtDisplay.Text = "0"
End Sub

Private Sub CommandButton19_Click()

End Sub

Private Sub CommandButton2_Click()
If InStr(txtDisplay.Text, ".") = 0 Then
txtDisplay.Text = txtDisplay.Text + "."
End If

End Sub

Private Sub CommandButton3_Click()
firstnum = txtDisplay.Text
txtDisplay.Text = ""
operator = "/"
End Sub

Private Sub CommandButton4_Click()
secondnum = txtDisplay.Text
If operator = "+" Then
answer = firstnum + secondnum
txtDisplay.Text = answer
ElseIf operator = "-" Then
answer = firstnum - secondnum
txtDisplay.Text = answer
ElseIf operator = "*" Then
answer = firstnum * secondnum
txtDisplay.Text = answer
ElseIf operator = "/" Then
answer = firstnum / secondnum
txtDisplay.Text = answer
ElseIf operator = "%" Then
answer = firstnum Mod secondnum
txtDisplay.Text = answer
End If

End Sub

Private Sub CommandButton5_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "1"
Else
txtDisplay.Text = txtDisplay.Text + "1"
End If
End Sub

Private Sub CommandButton6_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "2"
Else
txtDisplay.Text = txtDisplay.Text + "2"
End If

End Sub

Private Sub CommandButton7_Click()
firstnum = txtDisplay.Text
txtDisplay.Text = ""
operator = "-"
End Sub

Private Sub CommandButton8_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "3"
Else
txtDisplay.Text = txtDisplay.Text + "3"
End If
End Sub

Private Sub CommandButton9_Click()
If txtDisplay.Text = "0" Then
txtDisplay.Text = "4"
Else
txtDisplay.Text = txtDisplay.Text + "4"
End If
End Sub

Private Sub TextBox12_Change()

End Sub


Private Sub Label14_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub txtPostCode_Change()
cmdRefresh.Enabled = True
cmdTotal.Enabled = True
cmdView.Enabled = True


End Sub

Private Sub UserForm_Initialize()

chkTax.Value = True
txtName.SetFocus
cmdRefresh.Enabled = False
cmdTotal.Enabled = False
cmdView.Enabled = False
cmdAddItem.Enabled = False


End Sub

