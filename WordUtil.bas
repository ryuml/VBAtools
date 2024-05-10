Attribute VB_Name = "WordUtil"
Option Explicit


' 指定のWordファイルが既に開かれているか確認
Function CheckIfWordFileIsOpen(ByVal FilePath As String) As Boolean
    ' 関連するオブジェクトの宣言
    Dim wordApp As Object
    Dim doc As Object
    Dim IsDocOpen As Boolean
    
    IsDocOpen = False   ' 初期化
    
    ' エラーが発生してもスクリプトの実行を続行
    On Error Resume Next
    ' 既存のWordアプリケーションがあるか確認
    Set wordApp = GetObject(, "Word.Application")
    ' エラーハンドリングを元に戻す
    On Error GoTo 0
    
    ' Wordアプリケーションが存在する場合
    If Not wordApp Is Nothing Then
        ' 開かれている全てのドキュメントに対して処理を実行
        For Each doc In wordApp.Documents
            If doc.Path & "\" & doc.Name = FilePath Then
                ' ファイルが開かれている場合はTrueを返し、関数を終了
                IsDocOpen = True
                ' 既に開かれている場合は処理を中止
                Exit For
                End
            End If
        Next doc
    Else
        ' Wordアプリケーションが実行されていない場合
    End If
    
    ' オブジェクトの解放
    Set doc = Nothing
    Set wordApp = Nothing
    
    ' 結果返却
    CheckIfWordFileIsOpen = IsDocOpen
End Function
