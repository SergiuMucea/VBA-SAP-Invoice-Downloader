Sub Li3InvoiceDownloader()
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim docNumbers As Variant
    Dim i As Integer
    Dim conn As Object
    Dim sess As Object
    Dim userInput As String
    Dim errorInvoices As String
    Dim noAttachmentInvoices As String

    ' Prompt user to enter document numbers
    userInput = InputBox("Enter the document numbers separated by commas (max 5):" & vbCrLf & "(e.g.: 1234, 1235, 1236)", "Document Numbers")
    
    ' Check if the user clicked Cancel
    If userInput = "" Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
    End If
    
    ' Remove any trailing commas
    If Right(userInput, 1) = "," Then
        userInput = Left(userInput, Len(userInput) - 1)
    End If
    
    docNumbers = Split(userInput, ",")
    
    ' Check if the number of invoices exceeds the limit
    If UBound(docNumbers) > 4 Then
        MsgBox "The OpenText Imaging Windows Viewer cannot have more than 5 windows open." & vbCrLf & _
        "Therefore you can only enter a maximum of 5 invoices." & vbCrLf & _
        "The Downloader will now close, please try again.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize SAP GUI scripting
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0
    
    ' Loop through SAP connections to find the correct system
    For Each conn In application.Children
        For Each sess In conn.Children
            If sess.Info.SystemName = "LI3" Then
                Set session = sess
                Exit For
            End If
        Next sess
        If Not session Is Nothing Then Exit For
    Next conn

    ' Check if session is initialized
    If session Is Nothing Then
        MsgBox "SAP session for system 'Li3' could not be initialized. Please ensure SAP GUI is running and you have the necessary permissions.", vbCritical
        Exit Sub
    End If

    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If

    ' Initialize errorInvoices and noAttachmentInvoices
    errorInvoices = ""
    noAttachmentInvoices = ""

    ' Show progress form
    ProgressBarForm.StatusLabel.Caption = "Working on opening the Li3 invoices. Please wait..."
    ProgressBarForm.ProgressLabel.Caption = "Processed 0 invoice(s) out of " & UBound(docNumbers) - LBound(docNumbers) + 1
    ProgressBarForm.Show vbModeless

    ' Loop through each document number
    For i = LBound(docNumbers) To UBound(docNumbers)
    
        ' Trim any leading or trailing spaces from the document number
        docNumbers(i) = Trim(docNumbers(i))
        
        ' Open VF03 transaction
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVF03"
        session.findById("wnd[0]").sendVKey 0

        ' Wait for the document number field to be ready
        Set docField = session.findById("wnd[0]/usr/ctxtVBRK-VBELN")

        If docField Is Nothing Then
            MsgBox "Document number field not found. Please check the SAP layout.", vbCritical
            Exit Sub
        End If

        ' Enter document number
        docField.Text = docNumbers(i)
        session.findById("wnd[0]").sendVKey 0

        ' Check for error message
        If session.findById("wnd[0]/sbar").Text <> "" Then
            errorInvoices = errorInvoices & docNumbers(i) & vbCrLf
            GoTo NextDocument
        End If

        ' Open attachment
        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
        Set attachmentShell = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
        If attachmentShell Is Nothing Then
            noAttachmentInvoices = noAttachmentInvoices & docNumbers(i) & vbCrLf
            GoTo NextDocument
        End If
        attachmentShell.currentCellColumn = "BITM_DESCR"
        attachmentShell.selectedRows = "0"
        attachmentShell.contextMenu
        attachmentShell.doubleClickCurrentCell

        ' Close the attachment
        session.findById("wnd[1]").Close

NextDocument:
        ' Update progress label
        ProgressBarForm.ProgressLabel.Caption = (i + 1) & " invoice(s) out of " & UBound(docNumbers) - LBound(docNumbers) + 1
        DoEvents
    Next i

    ' Hide progress form
    Unload ProgressBarForm

    ' Inform the user about the operation result
    If errorInvoices = "" And noAttachmentInvoices = "" Then
        MsgBox "Operation completed successfully!", vbInformation
    Else
        Dim resultMessage As String
        resultMessage = "Operation completed with the following issues:" & vbCrLf
        If errorInvoices <> "" Then
            resultMessage = resultMessage & vbCrLf & "Incorrect invoice numbers:" & vbCrLf & errorInvoices
        End If
        If noAttachmentInvoices <> "" Then
            resultMessage = resultMessage & vbCrLf & "No attachment found for invoice numbers:" & vbCrLf & noAttachmentInvoices
        End If
        MsgBox resultMessage, vbExclamation
    End If
    
End Sub



