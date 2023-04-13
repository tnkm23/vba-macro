Attribute VB_Name = "Module1"
Sub UpdateIndex()
    ' 目次シートを取得する
    Dim wsTOC As Worksheet
    Set wsTOC = ThisWorkbook.Worksheets("目次")
     ' すべてのシート名を配列に追加する
    Dim sheet As Worksheet
    Dim arrSheetNames() As Variant
    Dim i As Integer
    i = 1
    For Each sheet In ThisWorkbook.Sheets
        ' 目次シートは除外する
        If sheet.Name <> "目次" Then
            ' 配列にシート名を追加する
            ReDim Preserve arrSheetNames(1 To i)
            arrSheetNames(i) = sheet.Name
            i = i + 1
        End If
    Next
     ' 目次シートのB列をクリアする
    wsTOC.Range("B3:B" & wsTOC.Cells(wsTOC.Rows.Count, "B").End(xlUp).Row).ClearContents
     ' 目次シートにシート名の見出しを追加する
    wsTOC.Range("B2").Value = "シート名"
     ' 目次シートにシート名を追加する
    wsTOC.Range("B3").Resize(UBound(arrSheetNames), 1).Value = Application.Transpose(arrSheetNames)
     ' 目次シートにリンクを追加する
    wsTOC.Range("B3").Resize(UBound(arrSheetNames), 1).Select
    For Each cell In Selection
        ' 各シート名に対するリンクを追加する
        ActiveSheet.Hyperlinks.Add cell, "", "'" & cell.Value & "'!A1"
    Next cell
End Sub
