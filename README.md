Macros:
in this we have make some visualization and it will show with the help of buttons.
Macro,Format,Highlight

VBA:First, we will learn to set up the visual basic editor using the developer tab, and insert a new module to start coding. As a first scenario, we'll create a custom function to calculate 
    discount percentages if a certain condition is true. Second, we'll create a sub procedure to clear the contents from a dataset in one click. We'll also add a message box and a button
    to confirm we want to clear the data. Finally, we'll learn to automate how to send an email from Excel containing a subject, a body, and the Excel file attached.
    Function Discount(Quantity, Price)

If Quantity > 25 Then
Discount = Quantity * Price * 0.2
Else
Discount = 0
End If

End Function
Sub ClearContent()
Answer = MsgBox("Confirm you want to clear?", vbYesNo)

If Answer = Yes Then
Rows("6:" & Rows.Count).ClearCounts
Else
Exit Sub
End If
End Sub

Sub Email()

Dim OutApp As Object
Dim OutMail As Object
Set OutApp = CreateObject("OutLook.Application")
Set OutMail = OutApp.CreateItem(0)

With OutMail

.To = "rajaniwalkunde@gmail.com"
.Subject = "Excel File"
.Body = "This is test"
.Attachments.Add ThisWorkbook.FullName
.Display

End With

Set OutApp = Nothing
Set OutMail = Nothing


End Sub

