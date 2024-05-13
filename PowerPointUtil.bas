Attribute VB_Name = "PowerPointUtil"
Option Explicit


' 指定のPowerPointファイルが既に開かれているか確認
Function CheckIfPwPtFileIsOpen(ByVal FilePath As String) As Boolean
    ' 関連するオブジェクトの宣言
    Dim ppApp As Object
    Dim pres As Object   ' Presentation
    Dim IsPresOpen As Boolean
    
    IsPresOpen = False   ' 初期化
    
    ' エラーが発生してもスクリプトの実行を続行
    On Error Resume Next
    ' 既存のPowerPointアプリケーションがあるか確認
    Set ppApp = GetObject(, "PowerPoint.Application")
    ' エラーハンドリングを元に戻す
    On Error GoTo 0
    
    ' PowerPointアプリケーションが存在する場合
    If Not ppApp Is Nothing Then
        ' 開かれている全てのプレゼンテーションに対して処理を実行
        For Each pres In ppApp.Presentations
            If pres.Path & "\" & pres.Name = FilePath Then
                ' ファイルが開かれている場合はTrueを返し、関数を終了
                IsPresOpen = True
                ' 既に開かれている場合は処理を中止
                Exit For
                End
            End If
        Next pres
    Else
        ' PowerPointアプリケーションが実行されていない場合
    End If
    
    ' オブジェクトの解放
    Set pres = Nothing
    Set ppApp = Nothing
    
    ' 結果返却
    CheckIfPwPtFileIsOpen = IsPresOpen
End Function



Sub GetPpAppHlinks(ByRef w_ws As Worksheet, ByRef objPres As Object, ByRef row_cnt As Integer)

' アクティブブックの全シートのセルと画像のハイパーリンクを抜き出し、新規シートに一覧で出力
    
    Dim ar()            As String       '// ハイパーリンク配列
    Dim hLink           As Hyperlink    '// ハイパーリンク(PowerPoint.Hyperlink)
'    Dim sCellAddress    As String       '// セル座標
    Dim sLinkAddress    As String       '// リンク先
    Dim sType           As String       '// 種類
    Dim s               As Variant      '// 配列の要素文字列
    Dim v               As Variant      '// 分割
'    Dim ppApp       As Object   'PowerPoint.Application
'    Dim p As Integer, y As Integer   'p:スライドページ Excel:y行
    Dim SldPage     As Integer
    Dim objShape    As Object 'PowerPoint.Shape 'パワポのシェイプ、テキスト、図形ほか
    Dim objSlide    As Object 'PowerPoint.Slide

'    On Error Resume Next  '取得エラー時に次へ
'    Set ppApp = GetObject(, "PowerPoint.Application")
'    On Error GoTo 0  'エラーを元に戻す※これを忘れると、デバッグ時にハマるから注意

    If ppApp Is Nothing Then
        MsgBox "パワーポイントを取得できません。"
        Exit Sub
    End If
    
    ReDim ar(0)

    '// アクティブプレゼンテーションの全スライドをループ
    For Each objSlide In objPres.Slides
        '// スライド内のハイパーリンクをループ
        For Each hLink In objSlide.Hyperlinks
            '// Range（セル）の場合
            If hLink.Type = msoHyperlinkRange Then
                sType = "Range"
            '// Shape（画像）の場合
            ElseIf hLink.Type = msoHyperlinkShape Then
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
            
            '// ファイル名＋スライド名＋スライド番号＋種類＋アドレス
            ar(UBound(ar)) = objPres.Name & vbTab & hLink.TextToDisplay & vbTab & objSlide.SlideNumber _
                                & vbTab & sType & vbTab & sLinkAddress
            ReDim Preserve ar(UBound(ar) + 1)
        Next
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
