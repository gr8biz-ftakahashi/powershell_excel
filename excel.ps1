# Excel開始処理
# Excelを操作する為の宣言
$excel = New-Object -ComObject Excel.Application

# 非表示処理
$excel.Visible = $False
# 高速化処理
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false

$new_excel_file_path = "C:\work\powershell\excel_new.xlsx"
$csv_file_path = "C:\work\powershell\enavi202203(1440).csv"


if (Test-Path $new_excel_file_path){
    # 既存のExcelファイルを開く
    $book = $excel.Workbooks.Open($new_excel_file_path)
}
else{
    # ワークブックを新規作成
    $book = $excel.Workbooks.Add()
}




# ワークシートを指定します
#$sheet = $excel.Worksheets.Item(1)
$sheet = $excel.Worksheets.Item("Sheet1")


# CSV読み込み
# 「QueryTableオブジェクト(＝クエリと接続)」を作成
$book.Worksheets.Add().Name = "tmp"
$tmp_sheet = $excel.Worksheets.Item("tmp")
$QueryTable = $tmp_sheet.QueryTables.Add("TEXT;$csv_file_path",$tmp_sheet.Range("A1"))
# 区切り文字に「カンマ区切り」を指定
$QueryTable.TextFileCommaDelimiter = $True
# $QueryTable.TextFileTabDelimiter  = $True
# $QueryTable.TextFileSemicolonDelimiter  = $True
# $QueryTable.TextFileSpaceDelimiter  = $True
# $QueryTable.TextFileOtherDelimiter   = ","

# 文字コード指定
#「Shift_JIS」を指定
#「UTF-8 」を指定
$QueryTable.TextFilePlatform = 65001
#$QueryTable.TextFilePlatform = 932

# ヘッダー行含める
$QueryTable.TextFileStartRow = 1
# ヘッダー行含めない
#$QueryTable.TextFileStartRow = 2

#
# 読み込むファイルの形式を指定
# 読み込むファイルの形式を【2:文字列】と指定するための配列を作成
$arrDataType = @()
for ($i=0; $i -lt 255; $i++){
    $arrDataType += 2
}
#$QueryTable.TextFileColumnDataTypes = $arrDataType[0..255]
# 読み込み実行
$QueryTable.Refresh($false)
# 名前を指定(後続処理で削除できるようにするため)
$QueryTable.Name = "仮テーブル"
# 作成された「QueryTableオブジェクト(＝クエリと接続)」を削除
$QueryTable.Delete()

# 最終行を取得
# https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xldirection?view=excel-pia
$max_row = $tmp_sheet.Range("A100").End(-4162).Row
$max_col = $tmp_sheet.Range("A1").End(-4161).Row
Write-Host($max_row ,$max_col )
$tmp_sheet.Range($tmp_sheet.Cells(1,1),$tmp_sheet.Cells($max_row,$max_col)).copy($sheet.Range("A5"))
$excel.Worksheets.Item("tmp").Delete()


# 指定したセルに計算式を入力
#$sheet.Range("C1").Value = "=A1+B2"
# セルの値をコピーして貼り付け
#$sheet.Range("C1").copy($sheet.Range("C2:C10"))





# 上記で作成されてしまう名前定義(仮テーブル)を削除
foreach($n in $book.Names){
    If ($n.Name -Like $loadSheetName + "!" + "仮テーブル*") {
        $n.Delete()
    }
}

# 表をテーブル化
#$sheet.ListObjects.Add(1,$sheet.Range($startRange).CurrentRegion,0,1).Name = $tableName



# 既に存在する場合は上書き保存
if (Test-Path $new_excel_file_path){
    # 上書き保存
    $book.Save()
}
else{
    # 名前をつけて保存
    $book.SaveAs($new_excel_file_path)
}


#高速化設定解除
$excel.DisplayAlerts = $true
$excel.ScreenUpdating = $true
$excel.EnableEvents = $true

# Excel終了処理
# Excelを閉じる
$excel.Quit()
# プロセスを解放する
$excel = $null
[GC]::Collect()
