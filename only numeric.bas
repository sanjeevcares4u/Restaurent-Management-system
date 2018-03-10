Attribute VB_Name = "Module1"
Public c As New ADODB.Connection
Public r As New ADODB.Recordset
Public s As String
Public sql As String

Public Sub conn()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
End Sub

Public Sub mdiform1exit()
Unload Entry_form
Unload Order
Unload Customer_Search
Unload Log_In
Unload Employee_Registration
Unload Employee_Search
Unload Employee_Attendence
Unload Salary_Entry
Unload Sal_Search
Unload Food_Entry
Unload View_Food
Unload Table_Status
Unload Create_user
Unload Salary_Setting
Unload Table_Entry
Unload Forgot_Password
Unload Report
Unload Password_Change
Unload About
Unload Order_search
End Sub

Public Sub mdiodisable()
MDIForm1.mnuorder.Enabled = False
MDIForm1.mnucustentry.Enabled = False
MDIForm1.mnucustsearch.Enabled = False
MDIForm1.mnuempentry.Enabled = False
MDIForm1.mnuempsearch.Enabled = False
MDIForm1.mnusalentry.Enabled = False
MDIForm1.mnuEmpatten.Enabled = False
MDIForm1.mnuSalsetting.Enabled = False
MDIForm1.mnuRegister.Enabled = False
MDIForm1.mnutblentry.Enabled = False
MDIForm1.mnuaddfood.Enabled = False
MDIForm1.mnuViewfood.Enabled = False
MDIForm1.mnutblstatus.Enabled = False
End Sub
'Option Explicit
Public Sub onlynum(tb As TextBox)
If Not IsNumeric(tb.Text) Then
        MsgBox "Please enter numeric value only.", vbInformation
      tb.Text = ""
        Cancel = True
        tb.SetFocus
End If
End Sub

Public Sub kypress(tb As TextBox)
If KeyAscii = 13 Or 32 Then
tb.SetFocus
End If
End Sub

