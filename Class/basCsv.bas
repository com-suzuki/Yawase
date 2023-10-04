Attribute VB_Name = "basCsv"
Option Explicit
'******************************************************************
'
'   プログラム名：ＣＳＶ操作関数
'   処理内容　　：
'   前提条件　　：
'   作成者　　　：株式会社 コム・エンジニアリング
'   作成年月日　：２００１／１２／１４
'   更新履歴　　：
'
'******************************************************************

'目　的　　：１レコードのＣＳＶデータをフィールド単位に変換する。
'条　件　　：
'結　果　　：
'引　数　　：(i)レコードデータ(String)
'　　　　　　(o)抽出したフィールドデータ（格納は配列）
'　　　　　　(o)抽出したフィールド数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００１／１２／１４
'更新履歴　：
'
Public Sub S_CSVtoTEXT(Recdata, FieldData() As Variant, FldCnt)

    Dim End_flg
    Dim fld_Data As String
    Dim Data_wk As Variant
    Dim RecLen
    
    Dim move_cnt    '移動位置
    Dim next_cnt    '次位置
    Dim entcnt      '現在位置
    
    Dim DQcnt As Integer 'ダブルクォーテーション数
    
    On Error Resume Next
    
    ReDim Preserve FieldData(0)
   
    End_flg = False
    FldCnt = 0
    entcnt = 0
    RecLen = Len(Recdata)
    Do While entcnt <= RecLen And (End_flg = False)
        If InStr(entcnt + 1, Recdata, Chr(34)) = (entcnt + 1) Then
            '先頭文字="
            FldCnt = FldCnt + 1
            ReDim Preserve FieldData(FldCnt)
            move_cnt = entcnt + 2
            Do While move_cnt <= RecLen
                '"と,検索
                next_cnt = InStr(move_cnt, Recdata, (Chr(34) & Chr(44)))
                If next_cnt = 0 Then
                    'なし
                    If InStr(RecLen, Recdata, Chr(34)) = RecLen Then
                        '後文字="
                        Data_wk = Mid(Recdata, entcnt + 2, RecLen - entcnt - 2)
                        S_CountDQ Data_wk, DQcnt    '"数算出
                        If (DQcnt Mod 2) = 0 Then
                            '偶数
                            S_CmbData Data_wk, fld_Data  'データ変換
                            Data_wk = fld_Data
                        Else
                            Data_wk = Mid(Recdata, entcnt + 1)   '残データ抽出
                        End If
                    Else
                        '後文字<>"
                        Data_wk = Mid(Recdata, entcnt + 1)   '残データ抽出
                    End If
                    FieldData(FldCnt) = Data_wk
                    End_flg = True
                    Exit Do
                Else
                    'あり
                    Data_wk = Mid(Recdata, entcnt + 2, next_cnt - entcnt - 2)
                    S_CountDQ Data_wk, DQcnt    '"数算出
                    If (DQcnt Mod 2) = 0 Then
                        '偶数
                        S_CmbData Data_wk, fld_Data  'データ変換
                        FieldData(FldCnt) = fld_Data
                        entcnt = next_cnt + 1
                        Exit Do
                    Else
                        '奇数
                        move_cnt = next_cnt + 2
                    End If
                End If
            Loop
        Else
            '先頭文字<>"
            FldCnt = FldCnt + 1
            ReDim Preserve FieldData(FldCnt)
            If InStr(entcnt + 1, Recdata, Chr(44)) = (entcnt + 1) Then
                '先頭文字=,
                FieldData(FldCnt) = ""
                entcnt = entcnt + 1
            Else
                '先頭文字<>","
                next_cnt = InStr(entcnt + 2, Recdata, Chr(44))
                '次のカンマあり？
                If next_cnt = 0 Then
                    'なし／残り全部結合
                    Data_wk = Mid(Recdata, entcnt + 1)
                    End_flg = True
                Else
                    'あり／終了
                    Data_wk = Mid(Recdata, entcnt + 1, next_cnt - entcnt - 1)
                    entcnt = next_cnt
                End If
                FieldData(FldCnt) = Data_wk
            End If
        End If
    Loop
    
End Sub

'目　的　　：ダブルクォーテーションの数を数える。
'条　件　　：
'結　果　　：
'引　数　　：(i)入力データ
'　　　　　　(o)ダブルクォーテーションの数
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００１／１２／１４
'更新履歴　：
'
Private Sub S_CountDQ(Data As Variant, cnt As Integer)
   
    Dim Pos As Integer
    Dim Pos_next As Integer
    Dim dLen As Integer
   
    On Error Resume Next
        
    dLen = Len(Data)
   
    cnt = 0
    Pos = 0
    Do While (Pos + 1) <= dLen
        Pos_next = InStr(Pos + 1, Data, Chr(34))
        If Pos_next <> 0 Then
            cnt = cnt + 1
            Pos = Pos_next
        Else
            'なし
            Exit Do
        End If
        Pos = Pos_next
    Loop
       
End Sub

'目　的　　：フィールドデータの変換
'条　件　　：
'結　果　　：
'引　数　　：(i)フィールドデータ
'　　　　　　(o)変換後のフィールドデータ
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング 渥美
'作成年月日：２００１／１２／１４
'更新履歴　：
'
Private Sub S_CmbData(Data As Variant, CmbData As String)
   
    Dim Pos As Integer
    Dim wData As String
    Dim Pos_next  As Integer
   
    On Error Resume Next
   
    '""->"置換
    Pos = 0
    Pos_next = InStr(Pos + 1, Data, (Chr(34) & Chr(34)))
    Do While True
        If Pos_next <> 0 Then
            wData = wData & Mid(Data, Pos + 1, Pos_next - Pos - 1) & Chr(34)
            Pos = Pos_next + 1
        Else
            wData = wData & Mid(Data, Pos + 1)
            Exit Do
        End If
        Pos_next = InStr(Pos + 1, Data, (Chr(34) & Chr(34)))
    Loop
   
    CmbData = wData
   
End Sub

