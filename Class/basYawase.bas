Attribute VB_Name = "basYawase"
Option Explicit

Private Const POSTAL_FILE_NAME = "Postal.CSV"       '�X�֔ԍ������p�t�@�C����
Private Const ODATE_FILE_NAME = "Odate.CSV"         '�J�Ó��p�t�@�C����

'�ځ@�I�@�@�F���喼�̂̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F objAdoSQL:clsAdoCore
'            strCode:����R�[�h
'            strOdate:�J�Ó�
'            strBcode:����R�[�h(�߂�l�p)
'�߂�l�@�@�F���喼��
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Get_Bname(objAdoSQL As clsAdoCore, strCode As String, strOdate As String, strBcode As String) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim adoRecordset2 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Global_Get_Bname_Err
    
    Global_Get_Bname = ""
    
    If Trim(strCode) = "" Then Exit Function
        
    strBcode = strCode
        
    '����ԍ��������甃��R�[�h���擾����
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
        
    '���Ӑ�}�X�^
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

    Call MsgBox("���喼�̂̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Bname_Err")

End Function

'�ځ@�I�@�@�F��Ж��̂̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F objAdoSQL:clsAdoCore
'�߂�l�@�@�F��Ж���
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Get_CompanyName(objAdoSQL As clsAdoCore) As String

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Global_Get_CompanyName_Err
    
    Global_Get_CompanyName = ""
        
    '�ݒ�}�X�^
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

    Call MsgBox("��Ж��̂̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_CompanyName_Err")

End Function

'�ځ@�I�@�@�F����Ōv�Z
'���@���@�@�F�ۂߒP�ʂ�0.01�E0.1�E1�E10�E100�E1000
'���@�ʁ@�@�F
'���@���@�@�FcurPrice:���z curRate:����ŗ� intBfraction�F�[���敪 curMarumeTani:�ۂߒP��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Get_Tax(ByVal curPrice As Currency, ByVal curRate As Currency, ByVal intBfraction As Integer, Optional curMarumeTani As Currency) As Currency

    Dim curRes As Currency

    On Error GoTo Global_Get_Tax_Err
    
    '����Ł����z�~�ŗ���100
    curRes = curPrice * curRate / 100
    
    '�����_�ȉ��̒[������
    curRes = Global_Fraction(curRes, intBfraction)
    
    '�ۂߒP�ʂ��ȗ����ꂽ�ꍇ
    If IsMissing(curMarumeTani) = True Then
        '�ۂߏ���
        Global_Get_Tax = Global_Rounding(curRes, intBfraction, 1)
    Else
        '�ۂߏ���
        Global_Get_Tax = Global_Rounding(curRes, intBfraction, curMarumeTani)
    End If
    
    Exit Function
    
Global_Get_Tax_Err:

    Call MsgBox("����Ōv�Z�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Tax_Err")

End Function

'�ځ@�I�@�@�F���勣���萔���v�Z
'���@���@�@�F�ۂߒP�ʂ�0.01�E0.1�E1�E10�E100�E1000
'���@�ʁ@�@�F
'���@���@�@�FcurPrice:���z curRate:����ŗ� intBfraction�F�[���敪 curMarumeTani:�ۂߒP��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O
'�쐬�N�����F�Q�O�P�P�^�O�W�^�P�X
'�X�V�����@�F
'
Public Function Global_Get_Brate(ByVal curPrice As Currency, ByVal curRate As Currency, ByVal intBfraction As Integer, Optional curMarumeTani As Currency) As Currency

    Dim curRes As Currency

    On Error GoTo Err
    
    '����Ł����z�~�����萔������100
    curRes = curPrice * curRate / 100
    
    '�����_�ȉ��̒[������
    curRes = Global_Fraction(curRes, intBfraction)
    
    '�ۂߒP�ʂ��ȗ����ꂽ�ꍇ
    If IsMissing(curMarumeTani) = True Then
        '�ۂߏ���
        Global_Get_Brate = Global_Rounding(curRes, intBfraction, 1)
    Else
        '�ۂߏ���
        Global_Get_Brate = Global_Rounding(curRes, intBfraction, curMarumeTani)
    End If
    
    Exit Function
    
Err:

    Call MsgBox("����Ōv�Z�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Tax_Err")

End Function

'�ځ@�I�@�@�F����ŗ��̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F objAdoSQL:clsAdoCore strOdate:�J�Ó�
'�߂�l�@�@�F����ŗ�
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Get_TaxRate(objAdoSQL As clsAdoCore, strOdate As String) As Currency

    Dim adoRecordset1 As New ADODB.Recordset
    Dim strSQL As String
    Dim curRate As Currency

    On Error GoTo Global_Global_Get_TaxRate_Err
    
    curRate = 0
        
    '����Ń}�X�^
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

    Call MsgBox("����ŗ��̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Global_Get_TaxRate_Err")

End Function

'�ځ@�I�@�@�F�ۂߏ���
'���@���@�@�F�ۂߒP�ʂ�0.01�E0.1�E1�E10�E100�E1000
'���@�ʁ@�@�F
'���@���@�@�F curPrice:���z intBfraction�F�[���敪 curMarumeTani:�ۂߒP��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Rounding(ByVal curPrice As Currency, ByVal intBfraction As Integer, ByVal curMarumeTani As Currency) As Currency

    Dim curRes As Currency
    Dim curMarume As Currency

    On Error GoTo Global_Rounding_Err
    
    '�ۂߒP��
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
    
    '�ۂߏ���
    curRes = curPrice * curMarume
    
    '�����_�ȉ��̒[������
    curRes = Global_Fraction(curRes, intBfraction)
    
    '�ۂ߂����ɖ߂�
    Global_Rounding = curRes / curMarume
    
    Exit Function
    
Global_Rounding_Err:

    Call MsgBox("�ۂߏ����G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Rounding_Err")

End Function

'�ځ@�I�@�@�F�����_�ȉ��̒[������
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F curPrice:���z intBfraction�F�[���敪
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�U�^�Q�V
'�X�V�����@�F
'
Public Function Global_Fraction(ByVal curPrice As Currency, ByVal intBfraction As Integer) As Currency

    Dim curRes As Currency

    On Error GoTo Global_Fraction_Err

    '�[������
    Select Case intBfraction
        Case FRACTION_OMISSION: '�؂�̂�
            curRes = Fix(curPrice)
        Case FRACTION_ROUNDOFF: '�l�̌ܓ�
            If curRes >= 0 Then
                curRes = Fix((curPrice + 0.5))
            Else
                curRes = Fix((curPrice - 0.5))
            End If
        Case FRACTION_ROUNDUP:  '�؂�グ
            If curRes >= 0 Then
                curRes = Fix((curPrice + 0.9999))
            Else
                curRes = Fix((curPrice - 0.9999))
            End If
    End Select

    Global_Fraction = curRes

    Exit Function

Global_Fraction_Err:

    Call MsgBox("�����_�ȉ��̒[�������G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Fraction_Err")

End Function

'�ځ@�I�@�@�F�A�ؗ��ʓ��������̏��v�^�C�g��
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F intDcode:���i�敪�̏\�̈�(�啪��)
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Public Function Global_CirculationTitle(intDcode As Integer) As String

    On Error GoTo Global_CirculationTitle_Err

    Global_CirculationTitle = ""

    Select Case intDcode
        Case 1:
            Global_CirculationTitle = "�j�t��"
        Case 2:
            Global_CirculationTitle = "��΍L�t��"
        Case 3:
            Global_CirculationTitle = "���t�L�t��"
        Case 4:
            Global_CirculationTitle = "���A�ʕ�"
        Case 5:
            Global_CirculationTitle = "�d����"
        Case Else
            Global_CirculationTitle = ""
    End Select

    Exit Function

Global_CirculationTitle_Err:

    Global_CirculationTitle = ""
    Call MsgBox("�A�ؗ��ʓ��������̏��v�^�C�g���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_CirculationTitle_Err")

End Function

'�ځ@�I�@�@�F�X�֔ԍ����猧���̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Public Function Global_Postal_to_Pref(strPath As String, strPostal As Variant) As Integer

    On Error GoTo Global_Postal_to_Pref_Err

    Global_Postal_to_Pref = 0
    
    If Not IsNull(strPostal) Or strPostal = "" Then
        Select Case Global_Get_Postal(strPath, Trim(CStr(strPostal)))
            Case "���m��":
                Global_Postal_to_Pref = 0 '����
            Case "���䌧", "�x�R��", "�ΐ쌧":
                Global_Postal_to_Pref = 1 '�k��
            Case "���s�{", "���{", "�ޗǌ�", "���Ɍ�", "���ꌧ", "�a�̎R��":
                Global_Postal_to_Pref = 2 '�֐�
            Case "�����s", "�_�ސ쌧", "��ʌ�", "��t��", "��錧", "�Ȗ،�", "�Q�n��":
                Global_Postal_to_Pref = 3 '�֓�
            Case Else
                Global_Postal_to_Pref = 4 '���̑�
        End Select
    Else
        Global_Postal_to_Pref = 4 '���̑�
    End If

    Exit Function

Global_Postal_to_Pref_Err:

    Global_Postal_to_Pref = ""
    Call MsgBox("�A�ؗ��ʓ��������̏��v�^�C�g���G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Postal_to_Pref_Err")

End Function

'�ځ@�I�@�@�F�X�֔ԍ�(��Q��)���猧���̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F strPostal:�X�֔ԍ��V��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�W�^�Q�Q
'�X�V�����@�F
'
Public Function Global_Get_Postal(strPath As String, strPostal As String) As String

    Dim intFileNum As Integer
    Dim strTextLine As String
    Dim strBuff() As String
    
    On Error GoTo Global_Get_Postal_Err

    Global_Get_Postal = ""

    '�t�@�C�����J��
    intFileNum = FreeFile
    Open strPath & "\" & POSTAL_FILE_NAME For Input Shared As #intFileNum
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strTextLine
        strBuff = Split(strTextLine, ",")
        '�X�֔ԍ��̏�Q�����猧����T��
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
    Call MsgBox("�X�֔ԍ����猧���̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_Postal_Err")

End Function

'�ځ@�I�@�@�F�O��J�Ó��̎擾
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O�@����
'�쐬�N�����F�Q�O�O�Q�^�O�X�^�O�S
'�X�V�����@�F
'
Public Function Global_Get_BeforeOdate(strPath As String, strOdate As String) As String

    Dim intFileNum As Integer
    Dim strTextLine As String
    Dim intIndex1 As Integer
    Dim varDate() As Variant
    
    On Error GoTo Global_Get_BeforeOdate_Err

    ReDim varDate(0)

    '�t�@�C�����J��
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
    Call MsgBox("�O��J�Ó��̎擾�G���[�I�I" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Global_Get_BeforeOdate_Err")

End Function



