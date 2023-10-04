Attribute VB_Name = "basYawase"
Option Explicit

Private Const POSTAL_FILE_NAME = "Postal.CSV"       '郵便番号検索用ファイル名
Private Const ODATE_FILE_NAME = "Odate.CSV"         '開催日用ファイル名

'目　的　　：買主名称の取得
'条　件　　：
'結　果　　：
'引　数　　： objAdoSQL:clsAdoCore
'            strCode:買主コード
'            strOdate:開催日
'            strBcode:買主コード(戻り値用)
'戻り値　　：買主名称
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Get_Bname(objAdoSQL As clsAdoCore, strCode As String, strOdate As String, strBcode As String) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Global_Get_Bname_Err
    
    Global_Get_Bname = ""
    
    If Trim(strCode) = "" Then Exit Function
        
    strBcode = strCode
        
    '買主番号履歴から買主コードを取得する
    strSQL = "SELECT * FROM MT071" & _
             " WHERE Bnum = " & strBcode & _
             " ORDER BY Sdate,Fdate,Bcode"
    adoRecordset1.Open strSQL, objAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    Do While Not adoRecordset1.EOF
        If Trim(adoRecordset1.Fields("Sdate")) <= strOdate And _
           Trim(adoRecordset1.Fields("Fdate")) >= strOdate Then
            strBcode = adoRecordset1.Fields("Bcode")
        End If
        adoRecordset1.MoveNext
    Loop
    adoRecordset1.Close
        
    '得意先マスタ
    strSQL = "{call sp_MT070;2(" & strBcode & ")}"
    adoRecordset2.Open strSQL, objAdoSQL.Connection, adOpenKeyset, adLockReadOnly
    If Not adoRecordset2.EOF Then
        Global_Get_Bname = IIf(IsNull(adoRecordset2.Fields("Bname")), "", Trim(adoRecordset2.Fields("Bname")))
    End If
    adoRecordset2.Close
    
    Set adoRecordset1 = Nothing
    Set adoRecordset2 = Nothing
    
    Exit Function
    
Global_Get_Bname_Err:

    Call MsgBox("買主名称の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Bname_Err")

End Function

'目　的　　：会社名称の取得
'条　件　　：
'結　果　　：
'引　数　　： objAdoSQL:clsAdoCore
'戻り値　　：会社名称
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Get_CompanyName(objAdoSQL As clsAdoCore) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Global_Get_CompanyName_Err
    
    Global_Get_CompanyName = ""
        
    '設定マスタ
    With adoRecordset1
        strSQL = "{call sp_MT010;1}"
        .Open strSQL, objAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        If Not .EOF Then
            Global_Get_CompanyName = IIf(IsNull(.Fields("Company")), "", Trim(.Fields("Company")))
        End If
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Exit Function
    
Global_Get_CompanyName_Err:

    Call MsgBox("会社名称の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_CompanyName_Err")

End Function

'目　的　　：消費税計算
'条　件　　：丸め単位は0.01・0.1・1・10・100・1000
'結　果　　：
'引　数　　：curPrice:金額 curRate:消費税率 intBfraction：端数区分 curMarumeTani:丸め単位
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Get_Tax(ByVal curPrice As Currency, ByVal curRate As Currency, ByVal intBfraction As Integer, Optional curMarumeTani As Currency) As Currency

    Dim curRes As Currency

    On Error GoTo Global_Get_Tax_Err
    
    '消費税＝金額×税率÷100
    curRes = curPrice * curRate / 100
    
    '小数点以下の端数処理
    curRes = Global_Fraction(curRes, intBfraction)
    
    '丸め単位が省略された場合
    If IsMissing(curMarumeTani) = True Then
        '丸め処理
        Global_Get_Tax = Global_Rounding(curRes, intBfraction, 1)
    Else
        '丸め処理
        Global_Get_Tax = Global_Rounding(curRes, intBfraction, curMarumeTani)
    End If
    
    Exit Function
    
Global_Get_Tax_Err:

    Call MsgBox("消費税計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Tax_Err")

End Function

'目　的　　：買主競売手数料計算
'条　件　　：丸め単位は0.01・0.1・1・10・100・1000
'結　果　　：
'引　数　　：curPrice:金額 curRate:消費税率 intBfraction：端数区分 curMarumeTani:丸め単位
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２０１１／０８／１９
'更新履歴　：
'
Public Function Global_Get_Brate(ByVal curPrice As Currency, ByVal curRate As Currency, ByVal intBfraction As Integer, Optional curMarumeTani As Currency) As Currency

    Dim curRes As Currency

    On Error GoTo Err
    
    '消費税＝金額×競売手数料率÷100
    curRes = curPrice * curRate / 100
    
    '小数点以下の端数処理
    curRes = Global_Fraction(curRes, intBfraction)
    
    '丸め単位が省略された場合
    If IsMissing(curMarumeTani) = True Then
        '丸め処理
        Global_Get_Brate = Global_Rounding(curRes, intBfraction, 1)
    Else
        '丸め処理
        Global_Get_Brate = Global_Rounding(curRes, intBfraction, curMarumeTani)
    End If
    
    Exit Function
    
Err:

    Call MsgBox("消費税計算エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Tax_Err")

End Function

'目　的　　：消費税率の取得
'条　件　　：
'結　果　　：
'引　数　　： objAdoSQL:clsAdoCore strOdate:開催日
'戻り値　　：消費税率
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Get_TaxRate(objAdoSQL As clsAdoCore, strOdate As String) As Currency

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim curRate As Currency

    On Error GoTo Global_Global_Get_TaxRate_Err
    
    curRate = 0
        
    '消費税マスタ
    With adoRecordset1
        strSQL = "{call sp_MT020;1}"
        .Open strSQL, objAdoSQL.Connection, adOpenKeyset, adLockReadOnly
        Do While Not .EOF
            If .Fields("Bdate") <= strOdate And strOdate <= .Fields("Fdate") Then
                curRate = .Fields("Tax")
                Exit Do
            End If
            .MoveNext
        Loop
        .Close
        Set adoRecordset1 = Nothing
    End With
    
    Global_Get_TaxRate = curRate
    
    Exit Function
    
Global_Global_Get_TaxRate_Err:

    Call MsgBox("消費税率の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Global_Get_TaxRate_Err")

End Function

'目　的　　：丸め処理
'条　件　　：丸め単位は0.01・0.1・1・10・100・1000
'結　果　　：
'引　数　　： curPrice:金額 intBfraction：端数区分 curMarumeTani:丸め単位
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Rounding(ByVal curPrice As Currency, ByVal intBfraction As Integer, ByVal curMarumeTani As Currency) As Currency

    Dim curRes As Currency
    Dim curMarume As Currency

    On Error GoTo Global_Rounding_Err
    
    '丸め単位
    Select Case curMarumeTani
        Case 0.01:
            curMarume = 100
        Case 0.1:
            curMarume = 10
        Case 1:
            curMarume = 1
        Case 10:
            curMarume = 0.1
        Case 100:
            curMarume = 0.01
        Case 1000:
            curMarume = 0.001
        Case Else
            curMarume = 1
    End Select
    
    '丸め処理
    curRes = curPrice * curMarume
    
    '小数点以下の端数処理
    curRes = Global_Fraction(curRes, intBfraction)
    
    '丸めを元に戻す
    Global_Rounding = curRes / curMarume
    
    Exit Function
    
Global_Rounding_Err:

    Call MsgBox("丸め処理エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Rounding_Err")

End Function

'目　的　　：小数点以下の端数処理
'条　件　　：
'結　果　　：
'引　数　　： curPrice:金額 intBfraction：端数区分
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０６／２７
'更新履歴　：
'
Public Function Global_Fraction(ByVal curPrice As Currency, ByVal intBfraction As Integer) As Currency

    Dim curRes As Currency

    On Error GoTo Global_Fraction_Err

    '端数処理
    Select Case intBfraction
        Case FRACTION_OMISSION: '切り捨て
            curRes = Fix(curPrice)
        Case FRACTION_ROUNDOFF: '四捨五入
            If curRes >= 0 Then
                curRes = Fix((curPrice + 0.5))
            Else
                curRes = Fix((curPrice - 0.5))
            End If
        Case FRACTION_ROUNDUP:  '切り上げ
            If curRes >= 0 Then
                curRes = Fix((curPrice + 0.9999))
            Else
                curRes = Fix((curPrice - 0.9999))
            End If
    End Select

    Global_Fraction = curRes

    Exit Function

Global_Fraction_Err:

    Call MsgBox("小数点以下の端数処理エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Fraction_Err")

End Function

'目　的　　：植木流通動向調査の小計タイトル
'条　件　　：
'結　果　　：
'引　数　　： intDcode:商品区分の十の位(大分類)
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２２
'更新履歴　：
'
Public Function Global_CirculationTitle(intDcode As Integer) As String

    On Error GoTo Global_CirculationTitle_Err

    Global_CirculationTitle = ""

    Select Case intDcode
        Case 1:
            Global_CirculationTitle = "針葉樹"
        Case 2:
            Global_CirculationTitle = "常緑広葉樹"
        Case 3:
            Global_CirculationTitle = "落葉広葉樹"
        Case 4:
            Global_CirculationTitle = "株、玉物"
        Case 5:
            Global_CirculationTitle = "仕立物"
        Case Else
            Global_CirculationTitle = ""
    End Select

    Exit Function

Global_CirculationTitle_Err:

    Global_CirculationTitle = ""
    Call MsgBox("植木流通動向調査の小計タイトルエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_CirculationTitle_Err")

End Function

'目　的　　：郵便番号から県名の取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２２
'更新履歴　：
'
Public Function Global_Postal_to_Pref(strPath As String, strPostal As Variant) As Integer

    On Error GoTo Global_Postal_to_Pref_Err

    Global_Postal_to_Pref = 0
    
    If Not IsNull(strPostal) Or strPostal = "" Then
        Select Case Global_Get_Postal(strPath, Trim(CStr(strPostal)))
            Case "愛知県":
                Global_Postal_to_Pref = 0 '県内
            Case "福井県", "富山県", "石川県":
                Global_Postal_to_Pref = 1 '北陸
            Case "京都府", "大阪府", "奈良県", "兵庫県", "滋賀県", "和歌山県":
                Global_Postal_to_Pref = 2 '関西
            Case "東京都", "神奈川県", "埼玉県", "千葉県", "茨城県", "栃木県", "群馬県":
                Global_Postal_to_Pref = 3 '関東
            Case Else
                Global_Postal_to_Pref = 4 'その他
        End Select
    Else
        Global_Postal_to_Pref = 4 'その他
    End If

    Exit Function

Global_Postal_to_Pref_Err:

    Global_Postal_to_Pref = ""
    Call MsgBox("植木流通動向調査の小計タイトルエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Postal_to_Pref_Err")

End Function

'目　的　　：郵便番号(上２桁)から県名の取得
'条　件　　：
'結　果　　：
'引　数　　： strPostal:郵便番号７桁
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０８／２２
'更新履歴　：
'
Public Function Global_Get_Postal(strPath As String, strPostal As String) As String

    Dim intFileNum As Integer
    Dim strTextLine As String
    Dim strBuff() As String
    
    On Error GoTo Global_Get_Postal_Err

    Global_Get_Postal = ""

    'ファイルを開く
    intFileNum = FreeFile
    Open strPath & "\" & POSTAL_FILE_NAME For Input Shared As #intFileNum
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strTextLine
        strBuff = Split(strTextLine, ",")
        '郵便番号の上２桁から県名を探す
        If Trim(strBuff(0)) = left$(strPostal, 2) Then
            Global_Get_Postal = Trim(strBuff(1))
            Exit Do
        End If
    Loop

    Close intFileNum

    Exit Function

Global_Get_Postal_Err:

    Close
    Global_Get_Postal = ""
    Call MsgBox("郵便番号から県名の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Postal_Err")

End Function

'目　的　　：前回開催日の取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０９／０４
'更新履歴　：
'
Public Function Global_Get_BeforeOdate(strPath As String, strOdate As String) As String

    Dim intFileNum As Integer
    Dim strTextLine As String
    Dim intIndex1 As Integer
    Dim varDate() As Variant
    
    On Error GoTo Global_Get_BeforeOdate_Err

    ReDim varDate(0)

    'ファイルを開く
    intFileNum = FreeFile
    Open strPath & "\" & ODATE_FILE_NAME For Input Shared As #intFileNum
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strTextLine
        ReDim Preserve varDate(UBound(varDate) + 1)
        varDate(UBound(varDate)) = strTextLine
    Loop
    Close intFileNum

    For intIndex1 = UBound(varDate) To 1 Step -1
        If varDate(intIndex1) < Mid$(strOdate, 6, 5) Then
            Global_Get_BeforeOdate = Mid$(strOdate, 1, 5) & CStr(varDate(intIndex1))
            Exit Function
        End If
        If intIndex1 = 1 Then
            Global_Get_BeforeOdate = CStr(CInt(Mid$(strOdate, 1, 4)) - 1) & "/" & varDate(UBound(varDate))
            Exit Function
        End If
    Next intIndex1

    Exit Function

Global_Get_BeforeOdate_Err:

    Close
    Global_Get_BeforeOdate = ""
    Call MsgBox("前回開催日の取得エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_BeforeOdate_Err")

End Function



