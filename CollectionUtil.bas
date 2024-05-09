Attribute VB_Name = "CollectionUtil"
Option Explicit

'===========================================================
'
' コレクション操作用モジュール
'
' [処理概要]
'　・コレクションのキー検索、アイテム検索
'
' [索引]
'  □ 1. ExistsKey（Collention内のキー検索関数）
'      ・第２引数をキーとしてItemメソッドを実行し、
' 　　 　結果をもとにキーの存在を確認する。
'  □ 2. ExistsItem（Collentionの格納データの存在チェック関数）
'      ・第２引数を検索対象として各メンバーと突合し、
' 　　 　メンバー内の存在チェック結果を返す。
'  □ 3. ExistsNoKeyItem（Keyが未指定版のExistsItem関数）
'      ・第２引数を検索対象として代入処理行い、
' 　　 　エラー結果を基に、存在チェック結果を返す。
'
'===========================================================

' モジュール名
Const MODULE_NAME = "CollectionUtil"




'*********************************************************
'* ExistsKey（Collention内のキー検索関数）
'*********************************************************
'* 第１引数 | Collection | 検索対象となるオブジェクト
'* 第２引数 |   String   | 検索するキー
'*  戻り値　|   Boolan   | True Or False ※False＠初期値
'*********************************************************
'*   説明   | 第２引数をキーとしてItemメソッドを実行し、
'*   　　   | 結果をもとにキーの存在を確認する。
'*********************************************************
'*   備考   | オブジェクト未設定の場合 ⇒ 戻り値「False」
'*   　　   | メンバー数「0」の場合 ⇒ 戻り値「False」
'*********************************************************
 
Function ExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '戻り値の初期値：False
    ExistsKey = False
     
    '変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Itemメソッドを実行
    Call objCol.Item(strKey)
         
    'エラー値がない場合：キー検索はヒット（戻り値：True）
    If Err.Number = 0 Then ExistsKey = True
 
End Function



'*********************************************************
'* ExistsItem（Collentionの格納データの存在チェック関数）
'*********************************************************
'* 第１引数 | Collection | 検索対象となるオブジェクト
'* 第２引数 |  Variant   | 検索するデータ
'*　戻り値  |   Boolan   | True Or False ※False＠初期値
'*********************************************************
'*   説明   | 第２引数を検索対象として各メンバーと突合し、
'*   　　   | メンバー内の存在チェック結果を返す。
'*********************************************************
'*   備考   | オブジェクト未設定の場合 ⇒ 戻り値「False」
'*   　　   | メンバー数「0」の場合 ⇒ 戻り値「False」
'*********************************************************
 
Function ExistsItem(objCol As Collection, varItem As Variant) As Boolean
     
    Dim v As Variant
     
    '戻り値の初期値：False
    ExistsItem = False
     
    '変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
     
    'Collectionの各メンバーと突合
    For Each v In objCol
         
        '突合結果が一致した場合：戻り値「True」にループ抜け
        If v = varItem Then ExistsItem = True: Exit For
         
    Next
     
End Function



'*********************************************************
'* ExistsNoKeyItem（Keyが未指定(インデックス形式)の
'  Collentionの格納データの存在チェック関数）
'*********************************************************
'* 第１引数 | Collection | 検索対象となるオブジェクト
'* 第２引数 |  Variant   | 検索するデータ
'*　戻り値  |   Boolan   | True Or False ※False＠初期値
'*********************************************************
'*   説明   | 第２引数を検索対象として代入処理行い、
'*   　　   | エラー結果を基に、存在チェック結果を返す。
'*********************************************************
'*   備考   | オブジェクト未設定の場合 ⇒ 戻り値「False」
'*   　　   | メンバー数「0」の場合 ⇒ 戻り値「False」
'*********************************************************
 
Function ExistsNoKeyItem(objCol As Collection, varItem As Variant) As Boolean
     
    Dim v As Variant
     
    '戻り値の初期値：False
    ExistsNoKeyItem = False
     
    '変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
    
    On Error Resume Next
    
    Set v = objCol(varItem)
    
    'エラー値がない場合：ワークシート存在
    If Err.Number = 0 Then ExistsNoKeyItem = True
    
End Function
