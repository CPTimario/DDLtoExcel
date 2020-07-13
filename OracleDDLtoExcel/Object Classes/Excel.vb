Imports System.Runtime.InteropServices
Imports XL = Microsoft.Office.Interop.Excel

Class Excel
    Private Const HEADER_ROW_START As Integer = 4
    Private Const COL_START As Integer = 2

    Private Title As String
    Private Tables As List(Of Table)

    Public Sub New(ByVal tables As List(Of Table), ByVal title As String)
        Me.Title = title
        Me.Tables = tables
    End Sub

    Public Sub Export()
        Dim xlApplication As XL.Application
        Dim xlWorkbooks As XL.Workbooks
        Dim xlWorkbook As XL.Workbook
        Dim xlWorksheet As XL.Worksheet
        Dim tableCounter As Integer = 0

        xlApplication = CreateObject("Excel.Application")
        xlApplication.Visible = True

        xlWorkbooks = xlApplication.Workbooks
        xlWorkbook = xlWorkbooks.Add
        xlWorksheet = xlWorkbook.ActiveSheet

        Call CreateSummarySheet(xlWorksheet, Title)

        Call ShowStatus(CREATING_SUMMARY_SHEET, tableCounter, Tables.Count, New String() {tableCounter.ToString, Tables.Count.ToString})

        For Each table As Table In Tables
            Application.DoEvents()
            If CancelFlg Then Exit Sub

            xlWorksheet = xlWorkbook.Worksheets.Add(, xlWorksheet)
            Call CreateSheetForTable(xlWorksheet, table)

            tableCounter += 1
            Call ShowStatus(CREATING_TABLE_SHEET, tableCounter, New String() {table.Name, tableCounter.ToString, Tables.Count.ToString})
        Next

        'xlApplication.Visible = True

        'Release COM Objects
        Marshal.FinalReleaseComObject(xlWorksheet)
        Marshal.FinalReleaseComObject(xlWorkbook)
        Marshal.FinalReleaseComObject(xlWorkbooks)
        Marshal.FinalReleaseComObject(xlApplication)
        xlWorksheet = Nothing
        xlWorkbook = Nothing
        xlWorkbooks = Nothing
        xlApplication = Nothing
    End Sub

    Private Sub CreateSummarySheet(ByRef xlWorksheet As XL.Worksheet, ByVal title As String)
        Dim xlHyperlinks As XL.Hyperlinks
        Dim xlRange As XL.Range
        Dim xlRange2 As XL.Range
        Dim xlFont As XL.Font
        Dim xlInterior As XL.Interior
        Dim xlBorders As XL.Borders
        Dim rowIndex As Integer = 0

        'Sheet Name
        xlWorksheet.Name = "SUMMARY"

        'Sheet Title
        xlRange = xlWorksheet.Range("A1")
        xlRange.Value = title & " SUMMARY"
        xlFont = xlRange.Font
        xlFont.Size = 14
        xlFont.Bold = True

        'Headers
        xlRange = xlWorksheet.Range("B3")
        xlRange.Value = "NO."

        xlRange = xlWorksheet.Range("C3")
        xlRange.Value = "TABLE NAME"

        xlRange = xlWorksheet.Range("D3")
        xlRange.Value = "DESCRIPTION"

        xlRange = xlWorksheet.Range("B3", "D3")
        xlFont = xlRange.Font
        xlFont.Size = 12
        xlFont.Bold = True
        xlInterior = xlRange.Interior
        xlInterior.Color = Color.DeepSkyBlue
        xlRange.HorizontalAlignment = XL.XlHAlign.xlHAlignCenter
        xlRange.VerticalAlignment = XL.XlVAlign.xlVAlignCenter

        Call ShowStatus(CREATING_SUMMARY_SHEET, rowIndex, Tables.Count, New String() {rowIndex.ToString, Tables.Count.ToString})

        xlHyperlinks = xlWorksheet.Hyperlinks
        For Each table As Table In Tables
            Application.DoEvents()
            If CancelFlg Then Exit Sub

            'NO.
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex, COL_START)
            xlRange.Value = Tables.IndexOf(table) + 1

            'TABLE NAME
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex, COL_START + 1)
            xlRange.Value = table.Name

            'Hyperlink
            xlHyperlinks.Add(xlRange, String.Empty, "'" & table.Name & "'!A1")

            rowIndex += 1
            Call ShowStatus(CREATING_SUMMARY_SHEET, rowIndex, New String() {rowIndex.ToString, Tables.Count.ToString})
        Next

        'Borders
        xlRange = xlWorksheet.Range("B3")
        xlRange2 = xlWorksheet.Cells(HEADER_ROW_START + rowIndex - 1, COL_START + 2)
        xlRange = xlWorksheet.Range(xlRange, xlRange2)
        xlBorders = xlRange.Borders
        xlBorders.LineStyle = XL.XlLineStyle.xlContinuous

        'Auto fit
        xlRange = xlWorksheet.Range("B1", "D1")
        xlRange = xlRange.EntireColumn
        xlRange.AutoFit()

        'Release COM Objects
        Marshal.FinalReleaseComObject(xlBorders)
        Marshal.FinalReleaseComObject(xlInterior)
        Marshal.FinalReleaseComObject(xlFont)
        Marshal.FinalReleaseComObject(xlRange2)
        Marshal.FinalReleaseComObject(xlRange)
        Marshal.FinalReleaseComObject(xlHyperlinks)
        xlBorders = Nothing
        xlInterior = Nothing
        xlFont = Nothing
        xlRange2 = Nothing
        xlRange = Nothing
        xlHyperlinks = Nothing
    End Sub

    Private Sub CreateSheetForTable(ByRef xlWorksheet As XL.Worksheet, ByVal table As Table)
        Dim xlHyperlinks As XL.Hyperlinks
        Dim xlRange As XL.Range
        Dim xlRange2 As XL.Range
        Dim xlFont As XL.Font
        Dim xlInterior As XL.Interior
        Dim xlBorders As XL.Borders
        Dim rowIndex As Integer = 0
        Dim refTable As Table
        Dim refColumn As Column
        Dim subAddress As String

        'Sheet Name
        xlWorksheet.Name = table.Name

        'Sheet Title
        xlRange = xlWorksheet.Range("A1")
        xlRange.Value = table.Name
        xlFont = xlRange.Font
        xlFont.Size = 14
        xlFont.Bold = True

        'Hyperlink
        xlHyperlinks = xlWorksheet.Hyperlinks
        xlRange = xlWorksheet.Range("K1")
        xlRange.Value = "Back to Summary"
        subAddress = "'SUMMARY'!C" & HEADER_ROW_START + Tables.IndexOf(table)
        xlHyperlinks.Add(xlRange, String.Empty, subAddress)

        'Headers
        xlRange = xlWorksheet.Range("B3")
        xlRange.Value = "NO."
        xlRange = xlWorksheet.Range("B3", "B4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("C3")
        xlRange.Value = "COLUMN NAME"
        xlRange = xlWorksheet.Range("C3", "C4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("D3")
        xlRange.Value = "DATA TYPE"
        xlRange = xlWorksheet.Range("D3", "D4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("E3")
        xlRange.Value = "CONSTRAINT"
        xlRange = xlWorksheet.Range("E3", "I3")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("E4")
        xlRange.Value = "NOT NULL"

        xlRange = xlWorksheet.Range("F4")
        xlRange.Value = "PRIMARY KEY"

        xlRange = xlWorksheet.Range("G4")
        xlRange.Value = "UNIQUE"

        xlRange = xlWorksheet.Range("H4")
        xlRange.Value = "FOREIGN KEY"

        xlRange = xlWorksheet.Range("I4")
        xlRange.Value = "CHECK"

        xlRange = xlWorksheet.Range("J3")
        xlRange.Value = "COMMENT"
        xlRange = xlWorksheet.Range("J3", "J4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("B3", "J4")
        xlFont = xlRange.Font
        xlFont.Size = 12
        xlFont.Bold = True
        xlInterior = xlRange.Interior
        xlInterior.Color = Color.LightGreen
        xlRange.HorizontalAlignment = XL.XlHAlign.xlHAlignCenter
        xlRange.VerticalAlignment = XL.XlVAlign.xlVAlignCenter

        For Each column As Column In table.Columns
            'NO.
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START)
            xlRange.Value = table.Columns.IndexOf(column) + 1

            'TABLE NAME
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 1)
            xlRange.Value = column.Name

            'DATA_TYPE
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 2)
            xlRange.Value = column.DataType.Type.EnumToString & column.DataType.Arguments

            'CONSTRAINTS
            For Each constraint As Constraint In column.Constraints
                xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + constraint.Type + 3)

                Select Case constraint.Type
                    Case Constraint._Type.NOT_NULL, Constraint._Type.PRIMARY_KEY, Constraint._Type.UNIQUE
                        xlRange.Value = "YES"
                    Case Constraint._Type.FOREIGN_KEY
                        refTable = Tables.Table(constraint.Reference.Key)
                        refColumn = refTable.Column(constraint.Reference.Value)
                        subAddress = "'" & constraint.Reference.Key & "'!C" & HEADER_ROW_START + refTable.Columns.IndexOf(refColumn) + 1
                        xlHyperlinks.Add(xlRange, String.Empty, subAddress)
                        xlRange.Value = constraint.Reference.Key & "." & constraint.Reference.Value
                    Case Constraint._Type.CHECK
                        xlRange.Value = constraint.Expression
                End Select
            Next

            'COMMENT
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 8)
            xlRange.Value = column.Comment

            rowIndex += 1
        Next

        'Borders
        xlRange = xlWorksheet.Range("B3")
        xlRange2 = xlWorksheet.Cells(HEADER_ROW_START + rowIndex, COL_START + 8)
        xlRange = xlWorksheet.Range(xlRange, xlRange2)
        xlBorders = xlRange.Borders
        xlBorders.LineStyle = XL.XlLineStyle.xlContinuous

        'Auto fit
        xlRange = xlWorksheet.Range("B1", "K1")
        xlRange = xlRange.EntireColumn
        xlRange.AutoFit()

        'Release COM Objects
        Marshal.FinalReleaseComObject(xlBorders)
        Marshal.FinalReleaseComObject(xlInterior)
        Marshal.FinalReleaseComObject(xlFont)
        Marshal.FinalReleaseComObject(xlRange2)
        Marshal.FinalReleaseComObject(xlRange)
        xlBorders = Nothing
        xlInterior = Nothing
        xlFont = Nothing
        xlRange2 = Nothing
        xlRange = Nothing
    End Sub
End Class
