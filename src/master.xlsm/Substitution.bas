Attribute VB_Name = "Substitution"
Private Const FILE_NAME As String = "origin.docx"
Private Const OUTPUT_NAME As String = "output"
Const COMPANY = "@company"
Const DATETIME = "@datetime"

Sub SubstitutionInWord()
    '���܂��Ȃ�
    Dim word As word.Application: Set word = New word.Application
    word.Visible = True
    'Word�t�@�C���̏ꏊ
    Dim docPath As String: docPath = ThisWorkbook.Path & "\" & FILE_NAME
    '�o�͐�̏ꏊ
    Dim outPath As String: outPath = ThisWorkbook.Path & "\" & OUTPUT_NAME
    '�㏑���ۑ��ɂ���
    Application.DisplayAlerts = False
    
    
    '�s����
    Dim rCompany As String
    Dim rDatetime As String
    Dim i
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        '�u���p������̓���(A���E��)
        rCompany = Worksheets(1).Range("A" & i)
        rDatetime = Worksheets(1).Range("E" & i).Text
        '�����pword
        Dim outputName As String: outputName = outPath & "\" & rCompany & ".docx"
        FileCopy docPath, outputName
        
        'Word�t�@�C���J��
        Dim doc As word.Document: Set doc = word.Documents.Open(Filename:=outputName, Visible:=True, ReadOnly:=False)
        '--- �u�������s���� ---'
        With doc.Content.Find
            '�Ж�
            .Text = COMPANY
            .Replacement.Text = rCompany
            .Execute Replace:=wdReplaceAll
            .ClearFormatting
            .Replacement.ClearFormatting
            '���t
            .Text = DATETIME
            .Replacement.Text = rDatetime
            .Execute Replace:=wdReplaceAll
        End With
        '�ۑ�
        doc.Save
         '�t�@�C�����鏈��
        doc.Close savechanges:=False
        Set doc = Nothing
    Next
    
    word.Quit
    Set word = Nothing
End Sub
