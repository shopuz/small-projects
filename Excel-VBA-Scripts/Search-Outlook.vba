Private m_blnWeOpenedExcel As Boolean
Function GetExcelWB() As Excel.Workbook
    Dim objExcel As Excel.Application
    On Error Resume Next
    m_blnWeOpenedExcel = False
    Set objExcel = GetObject(, "Excel.Application")
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.Application")
        m_blnWeOpenedExcel = True
    End If
    Set GetExcelWB = objExcel.Workbooks.Add
    Set objExcel = Nothing
End Function


Sub GetFromInbox()

    Dim olApp As Outlook.Application
    Dim olNs As NameSpace
    Dim Fldr As MAPIFolder
    Dim olMail As Variant
    Dim i As Integer
    Dim x As Integer
    Dim id As String
    Dim Email As String
    Dim agree As String
    Dim objWB As Excel.Workbook
    Dim objWS As Excel.Worksheet
    
    Set objWB = GetExcelWB()
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set Fldr = olNs.GetDefaultFolder(olFolderInbox)
    i = 2
    
    'initialize object workbook if not already
    If Not objWB Is Nothing Then
     Set objWS = objWB.Sheets(1)
     objWS.Cells(1, 1) = "Facebook_id"
     objWS.Cells(1, 2) = "Email"
     objWS.Cells(1, 3) = "Agree"
    End If
    
    'loop through each mail in the outlook's default folder
    For Each olMail In Fldr.Items
        'search for the position of facebook_id
        If InStr(olMail.Body, "facebook_id:") > 0 Then
            x = InStr(olMail.Body, "facebook_id:") + 15
            y = InStr(olMail.Body, "first_name")
            
            'id is in between the keywords facebook_id and first_name
            id = Mid(olMail.Body, x, y - x)
            
            
            x = InStr(olMail.Body, "email:") + 9
            
            If (InStr(olMail.Body, "I agree")) Then
                y = InStr(olMail.Body, "I agree")
                'email is in between the keywords email and I agree
                Email = Mid(olMail.Body, x, (y - x))
                agree = Mid(olMail.Body, y + 60, 3)
            Else
                Email = ""
                agree = ""
            End If
            
            
            ' code to fill a worksheet with data
             objWS.Cells(i, 1) = id
             objWS.Cells(i, 2) = Email
             objWS.Cells(i, 3) = agree
             objWS.Application.Visible = True
             objWS.Activate
          
            i = i + 1
        End If
    Next olMail
       
    'closing excel workbook
    objWB.Close SaveChanges:=True
    
    Set objWS = Nothing
    Set objWB = Nothing
    Set Fldr = Nothing
    Set olNs = Nothing
    Set olApp = Nothing

End Sub


