Attribute VB_Name = "IndiecraftButtons"
Option Explicit

' =====================================================================
'  INDIECRAFT - one-click buttons  (Insert invoice / Reset form / Download PDF)
'  Setup: import this file, then run  InstallButtons  once (Alt+F8).
' =====================================================================
Private Const FORM_SHEET As String = "New Order"
Private Const REG_SHEET As String = "Register"
Private Const GEN_SHEET As String = "Invoice Generator"
Private Const PREVIEW_ROW As Long = 43
Private Const STATUS_CELL As String = "C39"
Private Const REG_LAST_ROW As Long = 301
Private Const REG_COLS As Long = 42

' ---------- INSERT: save the invoice typed on the form into the Register ----------
Public Sub SaveInvoice()
    Dim wsF As Worksheet, wsR As Worksheet, wsG As Worksheet
    Dim nr As Long, c As Long, v As Variant, invNo As String
    Set wsF = ThisWorkbook.Worksheets(FORM_SHEET)
    Set wsR = ThisWorkbook.Worksheets(REG_SHEET)
    Set wsG = ThisWorkbook.Worksheets(GEN_SHEET)

    If Left$(CStr(wsF.Range(STATUS_CELL).Value), 5) <> "Ready" Then
        MsgBox "Cannot save yet: " & wsF.Range(STATUS_CELL).Value, vbExclamation, "Indiecraft"
        Exit Sub
    End If

    invNo = CStr(wsF.Cells(PREVIEW_ROW, 1).Value)
    nr = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row + 1
    If nr < 2 Then nr = 2
    If nr > REG_LAST_ROW Then
        MsgBox "The Register is full. Please ask for a bigger version.", vbExclamation, "Indiecraft"
        Exit Sub
    End If

    For c = 1 To REG_COLS
        v = wsF.Cells(PREVIEW_ROW, c).Value
        If Not IsError(v) Then
            If Len(CStr(v)) > 0 Then wsR.Cells(nr, c).Value = v
        End If
    Next c

    ResetForm
    wsG.Range("I4").Value = invNo
    wsG.Activate
    MsgBox "Invoice " & invNo & " saved." & vbCrLf & vbCrLf & "Now click DOWNLOAD PDF to save it as a PDF.", vbInformation, "Indiecraft"
End Sub

' ---------- RESET: clear the form ----------
Public Sub ResetForm()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets(FORM_SHEET)
    wsF.Range("C6:C19").ClearContents
    wsF.Range("B23:D30").ClearContents
End Sub

' ---------- DOWNLOAD PDF: save the invoice shown on the Invoice Generator ----------
Public Sub DownloadPDF()
    Dim wsG As Worksheet, folder As String, fname As String, nm As String, sep As String
    Set wsG = ThisWorkbook.Worksheets(GEN_SHEET)
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save this workbook first (File > Save As), then click again.", vbExclamation, "Indiecraft"
        Exit Sub
    End If
    If Len(Trim$(CStr(wsG.Range("E3").Value))) = 0 Then
        MsgBox "Nothing to export - pick an invoice first.", vbExclamation, "Indiecraft"
        Exit Sub
    End If
    nm = CleanName(CStr(wsG.Range("E3").Value) & " - " & CStr(wsG.Range("B4").Value))
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "Invoices PDF"
    On Error Resume Next
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    On Error GoTo 0
    fname = folder & sep & nm & ".pdf"
    wsG.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fname, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF saved:" & vbCrLf & fname, vbInformation, "Indiecraft"
End Sub

' ---------- PREVIEW + PDF of the form (before saving) ----------
Public Sub PreviewPDF()
    ThisWorkbook.Worksheets(GEN_SHEET).Range("I4").Value = "NEW (from form)"
    Application.Calculate
    DownloadPDF
End Sub

Private Function CleanName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "-")
    Next i
    CleanName = Trim$(s)
End Function

' ---------- run ONCE: draws the buttons on the sheets ----------
Public Sub InstallButtons()
    Dim ws As Worksheet, shp As Shape
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If Left$(shp.Name, 4) = "btn_" Then shp.Delete
        Next shp
    Next ws
    MakeButton ThisWorkbook.Worksheets(FORM_SHEET), "H5:J7", "INSERT INVOICE", "SaveInvoice", RGB(46, 125, 50)
    MakeButton ThisWorkbook.Worksheets(FORM_SHEET), "H9:J10", "RESET FORM", "ResetForm", RGB(96, 96, 96)
    MakeButton ThisWorkbook.Worksheets(FORM_SHEET), "H12:J13", "PREVIEW + PDF", "PreviewPDF", RGB(107, 30, 30)
    MakeButton ThisWorkbook.Worksheets(GEN_SHEET), "H12:I13", "DOWNLOAD PDF", "DownloadPDF", RGB(107, 30, 30)
    MsgBox "Buttons installed. Save the file (Ctrl+S).", vbInformation, "Indiecraft"
End Sub

Private Sub MakeButton(ws As Worksheet, addr As String, caption As String, macroName As String, clr As Long)
    Dim r As Range, s As Shape
    Set r = ws.Range(addr)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, r.Left, r.Top, r.Width, r.Height)
    s.Name = "btn_" & macroName
    s.Fill.ForeColor.RGB = clr
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.OnAction = macroName
    s.Placement = xlFreeFloating
End Sub
