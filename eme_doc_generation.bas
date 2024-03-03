Attribute VB_Name = "Module1"
Option Explicit

Sub CreateDocumentFromTemplate()
    Dim wdApp As Object ' Word.Application
    Dim wdDoc As Object ' Word.Document
    Dim xlApp As Object ' Excel.Application
    Dim xlDoc As Object ' Excel.Document
    Dim input_file As String
    Dim output_file As String
    Dim site_id, jv_id, state_id, site_name, rfnsa_id, engineer_name As String
    Dim dt_today As Date
    Dim cwd As String
    
    Dim cell As Range
    Dim file_name As String
    Dim ws As Worksheet
    Dim chkbox As OLEObject
    Dim shp As Shape
    
    Application.DisplayAlerts = False
    
    site_id = Range("A2").Value
    jv_id = Range("B2").Value
    site_name = Range("C2").Value
    rfnsa_id = Range("D2").Value
    engineer_name = Range("E2").Value
    dt_today = Date
    
    cwd = Application.ActiveWorkbook.Path
    
    Set ws = ThisWorkbook.Worksheets("EME DOC generator")
    For Each chkbox In ws.OLEObjects
        If chkbox.Object.Value = False Then
        GoTo Label1
        End If
        file_name = chkbox.Object.Caption
        ' Debug.Print (file_name)
        
    
        ' Set the paths to your template and output document
        input_file = cwd + "\" & file_name
        output_file = cwd + "\" & Replace(file_name, "site_id", (site_id))
        output_file = Replace(output_file, "site_name", (site_name))
    
        ' if filename with 'docx', then Create a new instance of Word Application
        If InStr(file_name, ".docx") Then
            Set wdApp = CreateObject("Word.Application")
            wdApp.Visible = True ' Set to True if you want to see the Word application
    
            ' Open the template document
            Set wdDoc = wdApp.Documents.Add(Template:=input_file, NewTemplate:=False, DocumentType:=0)
    
            ' Modify the content of the document as needed
            ' Replace placeholders with actual values, use parentheses to use argument
            Call ReplacePlaceholder(wdDoc, "rfnsa_id", (rfnsa_id))
            Call ReplacePlaceholder(wdDoc, "site_id", (site_id))
            Call ReplacePlaceholder(wdDoc, "site_name", (site_name))
            Call ReplacePlaceholder(wdDoc, "engineer_name", (engineer_name))
            Call ReplacePlaceholder(wdDoc, "create_date", (dt_today))
            ' Save the new document
            wdDoc.SaveAs2 output_file, 16 ' Change the path and format(16:'docx') as needed
    
    
            ' Close the template document without saving changes
            wdDoc.Close SaveChanges:=False
    
            ' Quit Word Application
            wdApp.Quit
            Set wdDoc = Nothing
            Set wdApp = Nothing
        End If
        
        If InStr(file_name, ".xlsx") Then
            ' Create a new instance of Excel Application
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Visible = True ' Set to True if you want to see the Word application
    
            ' Open the template document
            Set xlDoc = Workbooks.Add(input_file)
            'update site name, site id
            'xlDoc.Sheets("Mitigation").Range("B2").Value = site_name
            'xlDoc.Sheets("Mitigation").Range("E2").Value = site_id
            ' Save the new document
            
            ' if it's Form B file, replace jv_id to actual JV ID in filename
            ' replace JV_ID in/site name in Form A+ sheet
            If InStr(file_name, "jv_id") Then
                output_file = Replace(output_file, "jv_id", (jv_id))
                xlDoc.Sheets("Form A+").Range("G17").Value = site_name
                xlDoc.Sheets("Form A+").Range("Q17").Value = jv_id
                xlDoc.Sheets("Form A+").Range("G19").Value = site_id
                xlDoc.Sheets("Form A+").Range("Q19").Value = rfnsa_id
            End If
            
            ' EME analysis file replacement
            If InStr(file_name, "EME analysis") Then
                xlDoc.Sheets("Initial (A)").Range("B16").Value = site_name
                xlDoc.Sheets("Initial (A)").Range("A16").Value = jv_id
                xlDoc.Sheets("Initial (A)").Range("E16").Value = site_id
                
                Select Case Right(Left(jv_id, 2), 1)
                    Case "M"
                    state_id = "VIC"
                    
                    Case "S"
                    state_id = "NSW"
                    
                    Case "B"
                    state_id = "QLD"
                    
                    Case "A"
                    state_id = "SA"
                    
                    Case "P"
                    state_id = "WA"
                    
                    Case "C"
                    state_id = "ACT"
                    
                    Case "H"
                    state_id = "TAS"
                    
                    Case "D"
                    state_id = "NT"
                    
                End Select
                
                xlDoc.Sheets("Initial (A)").Range("C16").Value = state_id
            End If
            
            xlDoc.SaveAs output_file
            'SaveAs2 output_file ' Change the path and format(16:'docx') as needed
    
    
            ' Close the template document without saving changes
            xlDoc.Close SaveChanges:=False
    
            ' Quit Excel Application
            xlApp.Quit
            Set xlDoc = Nothing
            Set xlApp = Nothing
            'Set wdApp = CreateObject("Word.Application")
            'wdApp.Visible = True ' Set to True if you want to see the Word application
    
            ' Open the template document
            'Set wdDoc = wdApp.Documents.Add(Template:=input_file, NewTemplate:=False, DocumentType:=0)
    
            ' Modify the content of the document as needed
            ' Replace placeholders with actual values, use parentheses to use argument
            'Call ReplacePlaceholder(wdDoc, "rfnsa_id", (rfnsa_id))
            'Call ReplacePlaceholder(wdDoc, "site_id", (site_id))
            'Call ReplacePlaceholder(wdDoc, "site_name", (site_name))
            'Call ReplacePlaceholder(wdDoc, "engineer_name", (engineer_name))
            'Call ReplacePlaceholder(wdDoc, "create_date", (dt_today))
            ' Save the new document
            'wdDoc.SaveAs2 output_file, 16 ' Change the path and format(16:'docx') as needed
    
    
            ' Close the template document without saving changes
            'wdDoc.Close SaveChanges:=False
    
            ' Quit Word Application
            'wdApp.Quit
            'Set wdDoc = Nothing
            'Set wdApp = Nothing
        End If
Label1:
    Next chkbox
    
   Application.DisplayAlerts = True
    MsgBox ("Checklists, EME analysis and Form B Files for Site: " & site_id & " generated!")
End Sub



Function DownloadFileFromWeb(strURL As String, strSavePath As String) As Long
    ' strSavePath includes filename
    DownloadFileFromWeb = URLDownloadToFile(0, strURL, strSavePath, 0, 0)
End Function
Sub ReplacePlaceholder(doc As Object, placeholder As String, replacement As String)
    ' Replace text in content controls
    Dim cc As Object ' Word.ContentControl
    For Each cc In doc.ContentControls
        'Debug.Print cc.Title
        If cc.Title = placeholder Then
            cc.Range.Text = replacement
        End If
    Next cc
    
    ' Replace text in bookmarks
    Dim bm As Object ' Word.Bookmark
    For Each bm In doc.Bookmarks
        If bm.Name = placeholder Then
            bm.Range.Text = replacement
        End If
    Next bm
End Sub
