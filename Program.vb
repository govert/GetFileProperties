Imports ClosedXML.Excel

Module Program
    Sub Main(args As String())
        Dim wb As New XLWorkbook("perfDemo.xlsx")
        Dim ws As IXLWorksheet = wb.Worksheet("Sheet1")
        Dim msg As String

        Dim addrArray = {"a1:z500", "a1:z5000", "a1:z50000"}

        For Each addr In addrArray
            Dim rng As IXLRange = ws.Range(addr)

            Dim sw = Stopwatch.StartNew()
            Dim retBold = GetFontBold(rng)
            Dim cntBold As Integer
            For r = retBold.GetLowerBound(0) To retBold.GetUpperBound(0)
                For c = retBold.GetLowerBound(1) To retBold.GetUpperBound(1)
                    If retBold(r, c) Then cntBold += 1
                Next
                sw.Stop()
            Next
            msg = $"Count Bold for {addr} - Found {cntBold} in {sw.ElapsedMilliseconds} ms"
            Debug.Print(msg)
            Console.WriteLine(msg)

            sw = Stopwatch.StartNew()
            Dim retFml = GetHasFormula(rng)
            Dim cntFml = 0
            For r = retFml.GetLowerBound(0) To retFml.GetUpperBound(0)
                For c = retFml.GetLowerBound(1) To retFml.GetUpperBound(1)
                    If retFml(r, c) Then cntFml += 1
                Next
            Next
            sw.Stop()
            msg = $"Get Formulas for {addr} - Found {cntFml} in {sw.ElapsedMilliseconds} ms"
            Debug.Print(msg)
            Console.WriteLine(msg)

        Next

    End Sub

    Function GetHasFormula(rng As IXLRange) As Boolean(,)
        Dim cel As IXLCell
        Dim rows As Integer = rng.Rows.Count
        Dim cols As Integer = rng.Columns.Count
        Dim ret(rows, cols) As Boolean
        For r = 1 To rows
            For c = 1 To cols
                cel = rng.Cell(r, c)
                ret(r, c) = cel.HasFormula
            Next
        Next
        Return ret

    End Function

    Function GetFontBold(rng As IXLRange) As Boolean(,)
        Dim cnt = 0
        Dim cel As IXLCell
        Dim rows = rng.Rows.Count
        Dim cols = rng.Columns.Count
        Dim ret(rows, cols) As Boolean

        For r = 1 To rows
            For c = 1 To cols
                cel = rng.Cell(r, c)
                ret(r, c) = cel.Style.Font.Bold
            Next
        Next
        Return ret

    End Function
End Module
