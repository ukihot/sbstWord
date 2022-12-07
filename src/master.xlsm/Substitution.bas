Attribute VB_Name = "Substitution"
Private Const FILE_NAME As String = "origin.docx"
Private Const OUTPUT_NAME As String = "output"
Const COMPANY = "@company"
Const DATETIME = "@datetime"

Sub SubstitutionInWord()
    'おまじない
    Dim word As word.Application: Set word = New word.Application
    word.Visible = True
    'Wordファイルの場所
    Dim docPath As String: docPath = ThisWorkbook.Path & "\" & FILE_NAME
    '出力先の場所
    Dim outPath As String: outPath = ThisWorkbook.Path & "\" & OUTPUT_NAME
    '上書き保存にする
    Application.DisplayAlerts = False
    
    
    '行走査
    Dim rCompany As String
    Dim rDatetime As String
    Dim i
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        '置換用文字列の特定(A列とE列)
        rCompany = Worksheets(1).Range("A" & i)
        rDatetime = Worksheets(1).Range("E" & i).Text
        '生成用word
        Dim outputName As String: outputName = outPath & "\" & rCompany & ".docx"
        FileCopy docPath, outputName
        
        'Wordファイル開く
        Dim doc As word.Document: Set doc = word.Documents.Open(Filename:=outputName, Visible:=True, ReadOnly:=False)
        '--- 置換を実行する ---'
        With doc.Content.Find
            '社名
            .Text = COMPANY
            .Replacement.Text = rCompany
            .Execute Replace:=wdReplaceAll
            .ClearFormatting
            .Replacement.ClearFormatting
            '日付
            .Text = DATETIME
            .Replacement.Text = rDatetime
            .Execute Replace:=wdReplaceAll
        End With
        '保存
        doc.Save
         'ファイル閉じる処理
        doc.Close savechanges:=False
        Set doc = Nothing
    Next
    
    word.Quit
    Set word = Nothing
End Sub
