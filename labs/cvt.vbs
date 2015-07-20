Const dbg = False

Dim sh   ' シェル (カレントディレクトリ取得用)
Dim xl   ' Excel アプリ
Dim fso  ' File System Object
Dim file ' 出力ファイル

Dim ad   ' アドイン
Dim wb   ' 変換対象の表があるワークブック
Dim html ' 出力 HTML 格納用バッファ

' シェルを準備
set sh = CreateObject("WScript.Shell")

' Execl を開く
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel のウィンドウを表示するかどうか

' 出力ファイルオープン
Set fso  = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("test_table.xlsx.html", True)

' アドインを開く
Set ad = xl.Workbooks.Open(sh.CurrentDirectory & "\..\build\Excel2Html.xlam")

' 処理対象ワークブックを開く
Set wb = xl.Workbooks.Open(sh.CurrentDirectory & "\..\tools\test_table.xlsx")

' 変換対象セル選択 (TODO: 関数に引数で渡したほうがスタイリッシュ)
wb.Sheets("ver.2").Range("C4:N24").Select

' Excel 表 -> HTML へ変換
html = ad.Application.Run("Excel2Html.ConvertSelectedRangeToHtml")

' 結果をファイルへ出力
file.WriteLine html

MsgBox "test_table.xlsx.html に出力しました"

' 後始末
wb.Close
ad.Close
xl.Quit
