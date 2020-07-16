Imports System.Runtime.InteropServices
Imports OracleDDLtoExcel.SQL
Imports XL = Microsoft.Office.Interop.Excel

Class Excel
    Private Const HEADER_ROW_START As Integer = 4
    Private Const COL_START As Integer = 2

    Private Title As String
    Private Schemas As List(Of Schema)
    Private MessageArgs As String()

    Public Sub New(ByVal schemas As List(Of Schema), ByVal title As String)
        Me.Title = title
        Me.Schemas = schemas
    End Sub

    Public Sub Export()
        Dim xlApplication As XL.Application
        Dim xlWorkbooks As XL.Workbooks
        Dim xlWorkbook As XL.Workbook
        Dim xlWorksheet As XL.Worksheet
        Dim tableCounter As Integer
        Dim schemaString As String

        Call ShowStatus(CREATING_EXCEL_FILE)

        xlApplication = CreateObject("Excel.Application")
        xlApplication.Visible = False
        Application.DoEvents()

        xlWorkbooks = xlApplication.Workbooks
        xlWorkbook = xlWorkbooks.Add
        xlWorksheet = xlWorkbook.ActiveSheet

        If CancelFlg Then GoTo CloseExcel
        Call CreateSummarySheet(xlWorksheet, Title)

        For Each schema As Schema In Schemas
            If CancelFlg Then GoTo CloseExcel

            tableCounter = 0

            If schema.Name.Equals(String.Empty) Then
                schemaString = String.Empty
            Else
                schemaString = schema.Name & "."
            End If

            MessageArgs = New String() {String.Empty, String.Empty, tableCounter.ToString, schema.Tables.Count.ToString}
            Call ShowStatus(CREATING_TABLE_SHEET, tableCounter, schema.Tables.Count, MessageArgs)

            For Each table As Table In schema.Tables
                If CancelFlg Then GoTo CloseExcel

                tableCounter += 1
                MessageArgs = New String() {schemaString, table.Name, tableCounter.ToString, schema.Tables.Count.ToString}
                Call ShowStatus(CREATING_TABLE_SHEET, tableCounter, MessageArgs)

                xlWorksheet = xlWorkbook.Worksheets.Add(, xlWorksheet)
                Call CreateSheetForTable(xlWorksheet, schema, table)

                Application.DoEvents()
            Next
        Next

        xlApplication.Visible = True
        GoTo Release

CloseExcel:
        Call xlWorkbook.Close(False)
        Call xlApplication.Quit()

Release:
        'Release COM Objects
        Call ReleaseCOMObject(xlWorksheet)
        Call ReleaseCOMObject(xlWorkbook)
        Call ReleaseCOMObject(xlWorkbooks)
        Call ReleaseCOMObject(xlApplication)
    End Sub

    Private Sub CreateSummarySheet(ByRef xlWorksheet As XL.Worksheet, ByVal title As String)
        Dim xlRange As XL.Range
        Dim xlFont As XL.Font
        Dim rowIndex As Integer = 1

        'Sheet Name
        xlWorksheet.Name = "SUMMARY"

        'Sheet Title
        xlRange = xlWorksheet.Range("A" & rowIndex)
        xlRange.Value = title & " SUMMARY"
        xlFont = xlRange.Font
        xlFont.Size = 16
        xlFont.Bold = True

        For Each schema As Schema In Schemas
            rowIndex += 2
            Call PopulateSummarySheet(xlWorksheet, schema, rowIndex)
        Next

        'Auto fit
        xlRange = xlWorksheet.Range("B1", "D1")
        xlRange = xlRange.EntireColumn
        xlRange.AutoFit()

        'Release COM Objects
        Call ReleaseCOMObject(xlFont)
        Call ReleaseCOMObject(xlRange)
    End Sub

    Private Sub PopulateSummarySheet(ByVal xlWorksheet As XL.Worksheet, ByVal schema As Schema, ByRef rowIndex As Integer)
        Dim xlHyperlinks As XL.Hyperlinks
        Dim xlRange As XL.Range
        Dim xlRange2 As XL.Range
        Dim xlFont As XL.Font
        Dim xlBorders As XL.Borders
        Dim xlInterior As XL.Interior
        Dim headerStart As Integer
        Dim progress As Integer = 0
        Dim schemaString As String = String.Empty

        If Not schema.Name.Equals(String.Empty) Then
            schemaString = "for " & schema.Name & " schema "
        End If

        'Headers
        xlRange = xlWorksheet.Range("A" & rowIndex)
        xlRange.Value = schema.Name
        xlFont = xlRange.Font
        xlFont.Size = 14
        xlFont.Bold = True

        rowIndex += 1
        headerStart = rowIndex
        xlRange = xlWorksheet.Range("B" & rowIndex)
        xlRange.Value = "NO."

        xlRange = xlWorksheet.Range("C" & rowIndex)
        xlRange.Value = "TABLE NAME"

        xlRange = xlWorksheet.Range("D" & rowIndex)
        xlRange.Value = "DESCRIPTION"

        xlRange = xlWorksheet.Range("B" & headerStart, "D" & headerStart)
        xlFont = xlRange.Font
        xlFont.Size = 12
        xlFont.Bold = True
        xlInterior = xlRange.Interior
        xlInterior.Color = Color.DeepSkyBlue
        xlRange.HorizontalAlignment = XL.XlHAlign.xlHAlignCenter
        xlRange.VerticalAlignment = XL.XlVAlign.xlVAlignCenter

        MessageArgs = New String() {schemaString, progress.ToString, schema.Tables.Count.ToString}
        Call ShowStatus(CREATING_SUMMARY_SHEET, progress, schema.Tables.Count, MessageArgs)

        xlHyperlinks = xlWorksheet.Hyperlinks
        For Each table As Table In schema.Tables
            If CancelFlg Then Exit For

            rowIndex += 1
            progress += 1
            MessageArgs = New String() {schemaString, progress.ToString, schema.Tables.Count.ToString}
            Call ShowStatus(CREATING_SUMMARY_SHEET, progress, MessageArgs)

            'NO.
            xlRange = xlWorksheet.Cells(rowIndex, COL_START)
            xlRange.Value = schema.Tables.IndexOf(table) + 1

            'TABLE NAME
            xlRange = xlWorksheet.Cells(rowIndex, COL_START + 1)
            xlRange.Value = table.Name

            'Hyperlink
            If schema.Name.Equals(String.Empty) Then
                xlHyperlinks.Add(xlRange, String.Empty, "'" & table.Name & "'!A1")
            Else
                xlHyperlinks.Add(xlRange, String.Empty, "'" & schema.Name & "." & table.Name & "'!A1")
            End If

            Application.DoEvents()
        Next

        'Borders
        xlRange = xlWorksheet.Range("B" & headerStart)
        xlRange2 = xlWorksheet.Cells(rowIndex, COL_START + 2)
        xlRange = xlWorksheet.Range(xlRange, xlRange2)
        xlBorders = xlRange.Borders
        xlBorders.LineStyle = XL.XlLineStyle.xlContinuous

        'Release COM Objects
        Call ReleaseCOMObject(xlInterior)
        Call ReleaseCOMObject(xlBorders)
        Call ReleaseCOMObject(xlFont)
        Call ReleaseCOMObject(xlRange2)
        Call ReleaseCOMObject(xlRange)
        Call ReleaseCOMObject(xlHyperlinks)
    End Sub

    Private Sub CreateSheetForTable(ByRef xlWorksheet As XL.Worksheet, ByVal schema As Schema, ByVal table As Table)
        Dim xlHyperlinks As XL.Hyperlinks
        Dim xlRange As XL.Range
        Dim xlRange2 As XL.Range
        Dim xlFont As XL.Font
        Dim xlInterior As XL.Interior
        Dim xlBorders As XL.Borders
        Dim refTable As Table
        Dim refColumn As Column
        Dim subAddress As String
        Dim rowIndex As Integer = 0

        'Sheet Name
        If schema.Name.Equals(String.Empty) Then
            xlWorksheet.Name = table.Name
        Else
            xlWorksheet.Name = "(SCH" & (Schemas.IndexOf(schema) + 1) & ")." & table.Name
        End If

        'Sheet Title
        xlRange = xlWorksheet.Range("A1")
        If schema.Name.Equals(String.Empty) Then
            xlRange.Value = table.Name
        Else
            xlRange.Value = schema.Name & "." & table.Name
        End If
        xlFont = xlRange.Font
        xlFont.Size = 14
        xlFont.Bold = True

        'Hyperlink
        xlHyperlinks = xlWorksheet.Hyperlinks
        xlRange = xlWorksheet.Range("L1")
        xlRange.Value = "Back to Summary"
        subAddress = "'SUMMARY'!C" & GetTableRowIndex(table)
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
        xlRange.Value = "DEFAULT"
        xlRange = xlWorksheet.Range("E3", "E4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("F3")
        xlRange.Value = "CONSTRAINT"
        xlRange = xlWorksheet.Range("F3", "J3")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("F4")
        xlRange.Value = "NOT NULL"

        xlRange = xlWorksheet.Range("G4")
        xlRange.Value = "PRIMARY KEY"

        xlRange = xlWorksheet.Range("H4")
        xlRange.Value = "UNIQUE"

        xlRange = xlWorksheet.Range("I4")
        xlRange.Value = "FOREIGN KEY"

        xlRange = xlWorksheet.Range("J4")
        xlRange.Value = "CHECK"

        xlRange = xlWorksheet.Range("K3")
        xlRange.Value = "COMMENT"
        xlRange = xlWorksheet.Range("K3", "K4")
        xlRange.Merge()

        xlRange = xlWorksheet.Range("B3", "K4")
        xlFont = xlRange.Font
        xlFont.Size = 12
        xlFont.Bold = True
        xlInterior = xlRange.Interior
        xlInterior.Color = Color.LightGreen
        xlRange.HorizontalAlignment = XL.XlHAlign.xlHAlignCenter
        xlRange.VerticalAlignment = XL.XlVAlign.xlVAlignCenter

        For Each column As Column In table.Columns
            If CancelFlg Then Exit For

            'NO.
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START)
            xlRange.Value = table.Columns.IndexOf(column) + 1

            'TABLE NAME
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 1)
            xlRange.Value = column.Name

            'DATA_TYPE
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 2)
            xlRange.Value = column.DataType.Type.EnumToString & column.DataType.Arguments

            'DEFAULT
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 3)
            If Not column.DefaultValue.Equals(String.Empty) AndAlso column.DefaultValue.Chars(0).Equals(Chr(39)) Then
                xlRange.Value = Chr(39) & column.DefaultValue
            Else
                xlRange.Value = column.DefaultValue
            End If

            'CONSTRAINTS
            For Each constraint As Constraint In column.Constraints
                xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + constraint.Type + 4)

                Select Case constraint.Type
                    Case Constraint._Type.NOT_NULL, Constraint._Type.PRIMARY_KEY, Constraint._Type.UNIQUE
                        xlRange.Value = "YES"
                    Case Constraint._Type.FOREIGN_KEY
                        refColumn = constraint.ReferenceColumn
                        refTable = refColumn.ParentTable
                        If schema.Name.Equals(String.Empty) Then
                            subAddress = "'" & refTable.Name & "'!C" & GetColumnRowIndex(refColumn)
                        Else
                            subAddress = "'(SCH" & (Schemas.IndexOf(schema) + 1) & ")." & refTable.Name & "'!C" & GetColumnRowIndex(refColumn)
                        End If
                        xlHyperlinks.Add(xlRange, String.Empty, subAddress)
                        xlRange.Value = constraint.ReferenceColumn.ParentTable.Name & "." & constraint.ReferenceColumn.Name
                    Case Constraint._Type.CHECK
                        xlRange.Value = constraint.Expression
                End Select
            Next

            'COMMENT
            xlRange = xlWorksheet.Cells(HEADER_ROW_START + rowIndex + 1, COL_START + 9)
            xlRange.Value = column.Comment

            rowIndex += 1
            Application.DoEvents()
        Next

        'Borders
        xlRange = xlWorksheet.Range("B3")
        xlRange2 = xlWorksheet.Cells(HEADER_ROW_START + rowIndex, COL_START + 9)
        xlRange = xlWorksheet.Range(xlRange, xlRange2)
        xlBorders = xlRange.Borders
        xlBorders.LineStyle = XL.XlLineStyle.xlContinuous

        'Auto fit
        xlRange = xlWorksheet.Range("B1", "L1")
        xlRange = xlRange.EntireColumn
        xlRange.AutoFit()

        'Release COM Objects
        Call ReleaseCOMObject(xlBorders)
        Call ReleaseCOMObject(xlInterior)
        Call ReleaseCOMObject(xlFont)
        Call ReleaseCOMObject(xlRange2)
        Call ReleaseCOMObject(xlRange)
    End Sub

    Private Sub ReleaseCOMObject(ByVal comObject As Object)
        If Not IsNothing(comObject) Then
            Marshal.FinalReleaseComObject(comObject)
            comObject = Nothing
        End If
    End Sub

    Private Function GetTableRowIndex(ByVal table As Table) As Integer
        Dim schemaIndex As Integer = Schemas.IndexOf(table.ParentSchema)
        Dim tableIndex As Integer = table.ParentSchema.Tables.IndexOf(table)
        Dim rowIndex As Integer = HEADER_ROW_START + tableIndex + 1

        For index As Integer = 0 To schemaIndex - 1
            rowIndex += Schemas.Item(index).Tables.Count + 2
        Next

        Return rowIndex
    End Function

    Private Function GetColumnRowIndex(ByVal column As Column) As Integer
        Return column.ParentTable.Columns.IndexOf(column) + HEADER_ROW_START + 1
    End Function
End Class
