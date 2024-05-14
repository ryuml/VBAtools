Attribute VB_Name = "MyTools_1_0_0"
Option Explicit


Sub ActivateFirstSheet_WindowZoom100_SelectAnyCell()
'
' ActivateFirstSheet_WindowZoom100_SelectAnyCell Macro
' ブックの各シートの拡大率を100%に統一し、選択セルをA1、シートは先頭シートを選択した状態にします。
'

'
    Dim zoom_val As Integer, selected_cell As String
    zoom_val = 100
    selected_cell = "A1"
    
    Dim i As Long
    For i = 0 To Worksheets.Count - 1
        Worksheets(Worksheets.Count - i).Activate
        ActiveWindow.Zoom = zoom_val
        Range(selected_cell).Select
    Next i
End Sub



Sub Set_FontAndCellSize_inSheet_MeiryoUI()
'
' Set_FontAndCellSize_inSheet_MeiryoUI Macro
' シート内のすべてのセルのフォントとA～D列のセルのサイズを設定します。
'

'
    'アクティブシート上のすべてのフォントを設定
    Cells.Select
    With Selection.Font
        .Name = "Meiryo UI"
        '.Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    'アクティブシート上のすべての列の幅を設定
    Cells.Select
    Dim col_width_num As Double  '列の幅の値
    col_width_num = 8.1    'Meiryo UI の時のデフォルトの列の幅の値
    Selection.ColumnWidth = col_width_num
    
    'A列～D列の幅を設定
    'Meiryo UI の行の高さのデフォルト値(ポイント値) = 15ポイント = 25ピクセル = 1.8(列の幅の値)
    ' ↑ A列～D列のセルをすべて正方形にするための設定
    col_width_num = 1.8
    Columns("A:D").Select
    Selection.ColumnWidth = col_width_num
    
    'アクティブシート上のすべての行の高さを設定
    Cells.Select
    Dim point_num As Double  'ポイント値(行の高さの値)
    '列の幅の値(標準フォントの幅を基準とする特殊な単位)をポイント(VBAのExcel操作上の基本単位)へ変換
    point_num = Range("A1").Width
    Selection.RowHeight = point_num
    
    '終了処理
    Range("A1").Select
    
End Sub



Sub Copy_Sheet_x5()
'
' Copy_Sheet_x5 Macro
' 現在のシートのコピーを5シート分生成します。
'

'
    Dim sheet_name As String, sheet_num As Integer
    sheet_name = ActiveSheet.Name
    sheet_num = Sheets(sheet_name).Index
    
    '
    'Worksheets (sheet_num)
    '
    
    Dim i As Long
    For i = 1 To 5
        Sheets(sheet_name).Select
        Sheets(sheet_name).Copy After:=Sheets(i)
    Next i
    
End Sub



Sub Copy_Sheet_x10()
'
' Copy_Sheet_x10 Macro
' 現在のシートのコピーを10シート分生成します。
'

'
    Dim sheet_name As String, sheet_num
    sheet_name = ActiveSheet.Name
    sheet_num = Sheets(sheet_name).Index
    
    '
    'Worksheets (sheet_num)
    '
    
    Dim i As Long
    For i = 1 To 10
        Sheets(sheet_name).Select
        Sheets(sheet_name).Copy After:=Sheets(i)
    Next i
    
End Sub



Function ExistsSheet(sheetName As String) As Boolean
    Dim sh  As Worksheet

    '戻り値の初期値：False
    ExistsSheet = False

    On Error Resume Next

    Set sh = Worksheets(sheetName)

    'エラー値がない場合：ワークシート存在
    If Err.Number = 0 Then ExistsSheet = True

End Function



Sub Run_GetLinksDirAllFiles()
    Call MessageUserForm.Show(vbModal)
End Sub



Sub GetLinksDirAllFiles()
'// 指定のディレクトリ内のOfficeファイル内のリンクを取得

    Dim objItemWB       As Workbook
    Dim objWriteWS      As Worksheet
    Dim ppApp           As Object   '// PowerPointオブジェクト
    Dim FolderPath      As String
    Dim FilePath        As String
    Dim FileName        As String
    Dim ExcelExt1       As String
    Dim ExcelExt2       As String
    Dim WordExt1        As String
    Dim WordExt2        As String
    Dim PwPtExt1        As String
    Dim PwPtExt2        As String
    Dim RowCnt          As Integer
    Dim WriteWS_Name    As String
    Dim FileIsOpen      As Boolean
    Dim MsgText         As String
    Dim CntOpenFiles    As Integer
    
    '// --- 初期化
    MsgText = "処理候補ファイル／処理結果："
    CntOpenFiles = 0
    '// 拡張子指定
    ExcelExt1 = ".xls"
    ExcelExt2 = ".xlsx"
    WordExt1 = ".doc"
    WordExt2 = ".docx"
    PwPtExt1 = ".ppt"
    PwPtExt2 = ".pptx"
    '// 書き込みワークシート名設定
    WriteWS_Name = "LinkList"
    
    If ExistsSheet(WriteWS_Name) Then
        '// シートがある場合は内容をクリア
        Worksheets(WriteWS_Name).Activate
        ActiveSheet.Cells.Clear
    Else
        '// ブックの一番左に新規シートを追加
        Call Worksheets.Add(Before:=Worksheets(1))
        'Worksheets(1).Activate
        ActiveSheet.Name = WriteWS_Name   '// 名前変更
    End If
    
    '// マクロファイルのシートを設定
    Set objWriteWS = ThisWorkbook.ActiveSheet
    
    '// 行カウント初期化
    RowCnt = 1
    '// 列名設定
    objWriteWS.Cells(RowCnt, 1).Activate
    '// A列：ブック名
    ActiveCell.Value = "ファイル名"
    ActiveCell.ColumnWidth = 20
    ActiveCell.Font.Bold = True
    '// B列：シート名
    ActiveCell.Offset(0, 1).Value = "シート名/スライド名"
    ActiveCell.Offset(0, 1).ColumnWidth = 20
    ActiveCell.Offset(0, 1).Font.Bold = True
    '// C列：座標
    ActiveCell.Offset(0, 2).Value = "座標/スライド番号/ページ"
    ActiveCell.Offset(0, 2).ColumnWidth = 10
    ActiveCell.Offset(0, 2).Font.Bold = True
    '// D列：種類を出力
    ActiveCell.Offset(0, 3).Value = "種類"
    ActiveCell.Offset(0, 3).ColumnWidth = 10
    ActiveCell.Offset(0, 3).Font.Bold = True
    '// E列：アドレス
    ActiveCell.Offset(0, 4).Value = "アドレス"
    ActiveCell.Offset(0, 4).ColumnWidth = 20
    ActiveCell.Offset(0, 4).Font.Bold = True
    '// 行カウント更新
    RowCnt = RowCnt + 1
    
    '// フォルダ選択用ダイアログを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path '// 現在のフォルダパス
        If .Show = False Then Exit Sub
        FolderPath = .SelectedItems(1) '// フォルダパスを取得
    End With

    '// --- Excel処理
    '// 最初のファイル名を取得(Dir関数は"*.xls"でも"*.xls*"と同様の動作
    FileName = Dir(FolderPath & "\*" & ExcelExt1)
    Do While FileName <> ""
        FilePath = FolderPath & "\" & FileName
        
        '// ファイルが開いているか確認
        FileIsOpen = False   '// 初期化
        Set objItemWB = Nothing
        '// 自身のファイルと同じならフラグ立てる
        If FileName = ThisWorkbook.Name Then
            FileIsOpen = True
        Else
            '// FileIsOpenフラグ立っていなければ、他のブック一覧を確認
            For Each objItemWB In Workbooks
                If FileName = objItemWB.Name Then
                    FileIsOpen = True
                    Exit For
                End If
            Next
            '// オブジェクトの解放
            Set objItemWB = Nothing
        End If
        
        If FileIsOpen Then
            MsgText = MsgText & vbCrLf & FileName & ": " & "既に開いているためスキップしました"
        End If
        
        If Not FileIsOpen _
            And (LCase(FileName) Like ("*" & ExcelExt1) _
                Or LCase(FileName) Like ("*" & ExcelExt2)) Then
            '// フォームに読み込み中ファイル表示
            CntOpenFiles = CntOpenFiles + 1
            MsgText = MsgText & vbCrLf & FileName & ": " & "ファイルを読み込みました"
            
            '// ファイルを開く
            Workbooks.Open FileName:=FilePath
            
            '// 開いたファイルのリンクを調べて、書き込みシートへ記録
            Call ExcelExtractLinks(objWriteWS, Workbooks(FileName), RowCnt)
            '// 開いたファイルを閉じる
            'Application.DisplayAlerts = False
            Call Workbooks(FileName).Close(False)
            'Application.DisplayAlerts = True
        End If
        
        '// 更新処理
        FileName = Dir() '// 次のファイル名を取得
    Loop
    
    'CheckIfWordFileIsOpen (FilePath)
    
    
    '// --- PowerPoint処理
    '// 最初のファイル名を取得(Dir関数は"*.ppt"でも"*.ppt*"と同様の動作
    FileName = Dir(FolderPath & "\*" & PwPtExt1)
    Do While FileName <> ""
        FilePath = FolderPath & "\" & FileName
        
        '// ファイルが開いているか確認
        FileIsOpen = False   '// 初期化
        FileIsOpen = CheckIfPwPtFileIsOpen(FilePath)
        
        If FileIsOpen Then
            MsgText = MsgText & vbCrLf & FileName & ": " & "既に開いているためスキップしました"
        End If
        
        If Not FileIsOpen _
            And (LCase(FileName) Like ("*" & PwPtExt1) _
                Or LCase(FileName) Like ("*" & PwPtExt2)) Then
            '// フォームに読み込み中ファイル表示
            CntOpenFiles = CntOpenFiles + 1
            MsgText = MsgText & vbCrLf & FileName & ": " & "ファイルを読み込みました"
            
            '// ファイルを開く
            'Workbooks.Open FileName:=FilePath
            '// エラーが発生してもスクリプトの実行を続行
On Error Resume Next
            '// 既存のPowerPointアプリケーションがあるか確認
            Set ppApp = GetObject(, "PowerPoint.Application")
            '// エラーハンドリングを元に戻す
On Error GoTo 0
            '// PowerPointアプリケーションが存在する場合
            If Not ppApp Is Nothing Then
                '// 指定されたパスのファイルを開く
                Call ppApp.Presentations.Open(FileName:=FilePath)
            End If
            
            '// 開いたファイルのリンクを調べて、書き込みシートへ記録
            Call GetPpAppHlinks(objWriteWS, ppApp.Presentations(FileName), RowCnt)
            '// 開いたファイルを閉じる
            'Application.DisplayAlerts = False
            Call ppApp.Presentations(FileName).Close
            'Application.DisplayAlerts = True
            
            '// 使用後のオブジェクトを解放
            Set ppApp = Nothing
        End If
        
        '// 更新処理
        FileName = Dir() '// 次のファイル名を取得
    Loop
    
    
    '// Excelシートオブジェクトの解放
    Set objWriteWS = Nothing
    
    '// 終了処理
    If CntOpenFiles > 0 Then
        MsgText = MsgText & vbCrLf & vbCrLf & ">> ファイル読み込み数： " & CntOpenFiles
    End If
    MessageUserForm.MsgTextBox.Text = MsgText
    
End Sub



Sub ExcelExtractLinks(ByRef w_ws As Worksheet, ByRef r_wb As Workbook, _
                        ByRef row_cnt As Integer)
' アクティブブックの全シートのセルと画像のハイパーリンクを抜き出し、新規シートに一覧で出力
    
    Dim sht             As Worksheet    '// ワークシート
    Dim ar()            As String       '// ハイパーリンク配列
    Dim hLink           As Hyperlink    '// ハイパーリンク
    Dim sCellAddress    As String       '// セル座標
    Dim sLinkAddress    As String       '// リンク先
    Dim sType           As String       '// 種類
    Dim s               As Variant      '// 配列の要素文字列
    Dim v               As Variant      '// 分割
    Dim TmpUR           As Range        '// 一次的なUsedRangeのデータ範囲
    Dim CellDataRange   As Range        '// セルのデータ範囲複数
    Dim CellData        As Range        '// セルのデータ範囲単数
    
    ReDim ar(0)
    ReDim CellDataArr(0)
    
    '// アクティブブックの全シートをループ
    For Each sht In r_wb.Worksheets
        '// シート内のハイパーリンクをループ
        For Each hLink In sht.Hyperlinks
            '// Range（セル）の場合
            If hLink.Type = msoHyperlinkRange Then
                sCellAddress = hLink.Range.Address(False, False)
                sType = "Range"
            '// Shape（画像）の場合
            ElseIf hLink.Type = msoHyperlinkShape Then
                sCellAddress = hLink.Shape.TopLeftCell.Address(False, False)
                sType = "Shape"
            End If
            
            '// 外部リンクが設定されている場合
            If hLink.Address <> "" Then
                '// 外部へのハイパーリンクを取得
                sLinkAddress = hLink.Address
            '// 内部リンクが設定されている場合
            Else
                '// 内部へのハイパーリンクを取得
                sLinkAddress = hLink.SubAddress
            End If
            
            '// ブック名＋シート名＋セル座標＋種類＋アドレス
            ar(UBound(ar)) = r_wb.Name & vbTab & sht.Name & vbTab & sCellAddress _
                                & vbTab & sType & vbTab & sLinkAddress
            ReDim Preserve ar(UBound(ar) + 1)
        Next
        
On Error GoTo SkipNoFormula
        '// 範囲取得
        Set CellDataRange = Nothing   '// 初期化
        '// 何かしら書式・データが入っている範囲を取得
        Set TmpUR = sht.UsedRange
        If Not TmpUR Is Nothing Then
            For Each CellData In TmpUR
                If IsEmpty(CellData) Then
                    '// セル内容が空白（データが空 or 空文字列）しか無い場合はエラー回避のため数式判定せず、処理もしない設定
                    Set CellDataRange = Nothing
                Else
                    '// CellData のセル内容が空白（データが空 or 空文字列）でない場合、数式があるセル範囲を取得
                    Set CellDataRange = TmpUR.SpecialCells(xlCellTypeFormulas)
                End If
            Next
        End If
        '// オブジェクト解放
        Set TmpUR = Nothing
On Error GoTo 0
        
        '// HYPERLINKの値取得
        If Not CellDataRange Is Nothing Then
            '// CellDataRange に範囲（オブジェクト参照）がある場合
            For Each CellData In CellDataRange
                If CellData.Formula Like "*HYPERLINK*" Then
                    '// "HYPERLINK" が含まれている場合
                    sCellAddress = CellData.Address(False, False)
                    sType = "Formula"
                    sLinkAddress = CellData.Formula
                    sLinkAddress = Replace(sLinkAddress, "=", "", Count:=1)
                    
                    '// ブック名＋シート名＋セル座標＋種類＋アドレス
                    ar(UBound(ar)) _
                        = r_wb.Name & vbTab & sht.Name & vbTab & sCellAddress _
                            & vbTab & sType & vbTab & sLinkAddress
                    ReDim Preserve ar(UBound(ar) + 1)
                End If
            Next
        End If
        
SkipNoFormula:
        '// 数式セルが無かった場合スキップ
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next
    
    '// 配列(0)に格納済みの場合(ar(0)が空文字ではない)
    If ar(0) <> "" Then
        '// 余分な領域を削除
        ReDim Preserve ar(UBound(ar) - 1)
    End If
    
    '// ブックの一番左に新規シートを追加
    'Call Worksheets.Add(Before:=Worksheets(1))
    
    ' アクティブセル設定
    w_ws.Activate
    w_ws.Cells(row_cnt, 1).Activate
    
    '// ハイパーリンクの数ループ
    For Each s In ar
        If s <> "" Then
            '// TAB文字で分割
            v = Split(s, vbTab)
            '// A列にブック名を出力
            ActiveCell.Value = v(0)
            '// B列にシート名を出力
            ActiveCell.Offset(0, 1).Value = v(1)
            '// C列に座標を出力
            ActiveCell.Offset(0, 2).Value = v(2)
            '// D列に種類を出力
            ActiveCell.Offset(0, 3).Value = v(3)
            '// E列にアドレスを出力
            ActiveCell.Offset(0, 4).Value = v(4)
            
            '// 下のセルを選択
            row_cnt = row_cnt + 1
            ActiveCell.Offset(1, 0).Select
        End If
    Next
End Sub
