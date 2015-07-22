Const dbg = False

Dim sh          ' シェル (カレントディレクトリ取得用)
Dim xl          ' Excel アプリ
Dim fso         ' File System Object
Dim outFile     ' 出力ファイル
Dim outFileName ' 出力ファイル名
Dim inFileName  ' 入力ファイル名
Dim inSheetName ' 入力シート名
Dim inRange     ' 入力範囲
Dim arg

Dim ad   ' アドイン
Dim wb   ' 変換対象の表があるワークブック
Dim html ' 出力 HTML 格納用バッファ
Const example = "Example: cvt.vbs [..\tools\test_table.xlsx]ver.2!C4:N24"

' 引数チェック
If WScript.Arguments.Count = 0 Then
    MsgBox "Invalid arguments." & vbNewLine & _
           "Usage: cvt.vbs convert-range ..." & vbNewLine & _
           example
    WScript.Quit
End If

' シェルを準備
set sh = CreateObject("WScript.Shell")

' Execl を開く
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel のウィンドウを表示するかどうか

' アドインを開く
Set ad = xl.Workbooks.Open(sh.CurrentDirectory & "\..\build\Excel2Html.xlam")

For Each arg In WScript.Arguments
    Dim ok
    Dim pSt, pEn
    Dim exc

    WScript.echo "* ARG = " & arg
    ' 引数パース
    pSt = InStr(arg, "[")
    pEn = InStrRev(arg, "]")
    ok = False
    If pSt = 0 Or pEn = 0 Then
        WScript.echo "Error: Invalid convert-range format."
    ElseIf pSt + 1 = pEn Then
        WScript.echo "Error: No filename specified."
    Else
        inFileName = Mid(arg, pSt + 1, pEn - pSt - 1)
        ok = True
    End If

    If ok Then
        ' フルパス指定ではない時はカレントディレクトリを補う
        If Mid(inFileName, 2, 1) <> ":" And Left(inFileName, 2) <> "\\" Then
            inFileName = sh.CurrentDirectory & "\" & inFileName
        End If
        outFileName = inFileName & ".html"

        ' シート名取得
        exc = InStrRev(arg, "!")
        If exc = 0 Then
            WScript.echo "Error: Exclamation mark not found"
            ok = False
        ElseIf pEn + 1 = exc Then
            WScript.echo "Error: Sheet name not specified."
            ok = False
        End If
    End If

    If ok Then
        inSheetName = Mid(arg, pEn + 1, exc - pEn - 1)
        inRange = Mid(arg, exc + 1)
        If Len(inRange) = 0 Then
            WScript.echo "Error: Range not specified"
            ok = False
        End If
    End If

    If ok Then
        ' 処理対象ワークブックを開く
        WScript.echo "Opening " & inFileName & "..."
        Set wb = xl.Workbooks.Open(inFileName)

        ' 出力ファイルオープン
        Set fso  = CreateObject("Scripting.FileSystemObject")
        Set outFile = fso.CreateTextFile(outFileName, True)

        ' 変換対象セル選択 (TODO: 関数に引数で渡したほうがスタイリッシュ)
        wb.Sheets(inSheetName).Range(inRange).Select

        ' Excel 表 -> HTML へ変換
        WScript.echo "Converting..."
        html = ad.Application.Run("Excel2Html.ConvertSelectedRangeToHtml")

        ' 結果をファイルへ出力して閉じる
        outFile.WriteLine html
        outFile.Close

        ' ワークブックを閉じる
        wb.Close
        WScript.echo "Success." & vbNewLine
    End If
Next

' 後始末
ad.Close
xl.Quit
WScript.echo "Done."
