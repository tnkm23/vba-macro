Attribute VB_Name = "Module1"
Sub UpdateIndex()
    ' �ڎ��V�[�g���擾����
    Dim wsTOC As Worksheet
    Set wsTOC = ThisWorkbook.Worksheets("�ڎ�")
     ' ���ׂẴV�[�g����z��ɒǉ�����
    Dim sheet As Worksheet
    Dim arrSheetNames() As Variant
    Dim i As Integer
    i = 1
    For Each sheet In ThisWorkbook.Sheets
        ' �ڎ��V�[�g�͏��O����
        If sheet.Name <> "�ڎ�" Then
            ' �z��ɃV�[�g����ǉ�����
            ReDim Preserve arrSheetNames(1 To i)
            arrSheetNames(i) = sheet.Name
            i = i + 1
        End If
    Next
     ' �ڎ��V�[�g��B����N���A����
    wsTOC.Range("B3:B" & wsTOC.Cells(wsTOC.Rows.Count, "B").End(xlUp).Row).ClearContents
     ' �ڎ��V�[�g�ɃV�[�g���̌��o����ǉ�����
    wsTOC.Range("B2").Value = "�V�[�g��"
     ' �ڎ��V�[�g�ɃV�[�g����ǉ�����
    wsTOC.Range("B3").Resize(UBound(arrSheetNames), 1).Value = Application.Transpose(arrSheetNames)
     ' �ڎ��V�[�g�Ƀ����N��ǉ�����
    wsTOC.Range("B3").Resize(UBound(arrSheetNames), 1).Select
    For Each cell In Selection
        ' �e�V�[�g���ɑ΂��郊���N��ǉ�����
        ActiveSheet.Hyperlinks.Add cell, "", "'" & cell.Value & "'!A1"
    Next cell
End Sub
