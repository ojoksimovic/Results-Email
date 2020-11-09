Attribute VB_Name = "Module1"
Public Sub Emails()
Attribute Emails.VB_ProcData.VB_Invoke_Func = "E\n14"

Dim Splitcode As Range
Dim Department_Range As Range
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim Masterdata_Range As Range
Dim Masterdata As Range
Dim Emails As Range
Dim Email_Range As Range
Dim Department_Column As Integer
Dim Data_Row As Integer
Dim SendTo As String
Dim CCTo As String
Dim strSubject As String
Dim AcctMgrEmail As String
Dim OlApp As Object
Dim olItem As Outlook.MailItem
Dim Recip As Outlook.Recipient
Dim rng As Range
Dim YearOfAward As String



MsgBox ("There will be a series of questions asked now. If you're unsure about what to enter, go bug Olivera.")
Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

Set Masterdata_Range = Application.InputBox("Select the range of the entire results data table (including the header).", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Masterdata", _
            RefersTo:=Masterdata_Range
            
Set Department_Range = Application.InputBox("Select the entire range of departments. Do not include headers. Repeated Departments are OK.", _
Type:=8)

'Set Department_Range = Range("C2:C100")


Sheets.Add(After:=Sheets(Worksheet_Name)).Name = "Departments_VBA"
Department_Range.Copy Worksheets("Departments_VBA").Range("A1")

Worksheets("Departments_VBA").Range(Range("A1"), Range("A1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
ActiveWorkbook.Names.Add _
            Name:="Splitcode", _
            RefersTo:=Worksheets("Departments_VBA").Range(Range("A1"), Range("A1").End(xlDown))
            
'Set Email_Range = Application.InputBox("Select the range of the entire emails table.", _
Type:=8)
'Set Email_Range = Range("Email_Range")
'ActiveWorkbook.Names.Add _
            Name:="Emails", _
            RefersTo:=Email_Range
            
Workbooks(Workbook_Name).Sheets(Worksheet_Name).Select

            
'Department_Column = Application.InputBox("Which Column # contains the units? Enter the number only.", _
Type:=1)
'Data_Row = Application.InputBox("Which Row # does the data begin? This does NOT include the header.", _
Type:=1)

Department_Column = "3"
Data_Row = "2"



Workbooks(Workbook_Name).Sheets(Worksheet_Name).Select

For Each Cell In Range("Splitcode")

With Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Masterdata")
.AutoFilter Field:=Department_Column, Criteria1:=Cell.Value, Operator:=xlFilterValues
End With

SendTo1 = Application.VLookup(Cell.Value, Range("Emails"), 2, False)
SendTo2 = Application.VLookup(Cell.Value, Range("Emails"), 3, False)
YearOfAward = Cells(1, 15).Value
SendTo = SendTo1 & "; " & SendTo2
CCTo = Application.VLookup(Cell.Value, Range("Emails"), 4, False)
AcctMgrEmail = Application.VLookup(Cell.Value, Range("Emails"), 5, False)
strSubject = YearOfAward & " CIHR CGS Doctoral SGS Results - " & Cell.Value
' if adding attachment
'strAttachment = strAttachPath & xlSheet.Range("E" & rCount)

Set rng = Worksheets(Worksheet_Name).Range(Range("A1").End(xlDown), Range("A1").End(xlToRight))

Set OlApp = GetObject(, "Outlook.Application")

Worksheets(Worksheet_Name).Range("J1").Copy
Worksheets(Worksheet_Name).Range("K1").PasteSpecial Paste:=xlPasteValues


If Worksheets(Worksheet_Name).Range("K1").Value > 0 Then Set olItem = OlApp.CreateItemFromTemplate("C:\Users\olive\OneDrive - University of Toronto\VBA Instructions\Demo\Results Emails\CIHR\Successful.oft") Else: Set olItem = OlApp.CreateItemFromTemplate("C:\Users\olive\OneDrive - University of Toronto\VBA Instructions\Demo\Results Emails\CIHR\Unsuccessful.oft")

With olItem
.SentOnBehalfOfName = AcctMgrEmail
.To = SendTo
.CC = CCTo
.Subject = strSubject
.HTMLBody = Replace(.HTMLBody, "CIHR_Results", RangetoHTML(rng))


'if adding attachments:
'.Attachments.Add strAttachment

'.Save
.Display
'.Send

End With

Next Cell

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).AutoFilter.ShowAllData

Application.ActiveWorkbook.Names("Splitcode").Delete
Application.ActiveWorkbook.Names("Masterdata").Delete
Application.DisplayAlerts = False
Application.ActiveWorkbook.Worksheets("Departments_VBA").Delete
Application.DisplayAlerts = True

End Sub

Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim i As Integer
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    Application.ScreenUpdating = False
    
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .UsedRange.EntireColumn.AutoFit
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
        
        For i = 7 To 12
            With .UsedRange.Borders(i)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        Next i
        
    End With
 
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    TempWB.Close savechanges:=False

    Kill TempFile
    Application.ScreenUpdating = True
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function

