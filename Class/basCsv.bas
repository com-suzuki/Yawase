Attribute VB_Name = "basCsv"
Option Explicit
'******************************************************************
'
'   �v���O�������F�b�r�u����֐�
'   �������e�@�@�F
'   �O������@�@�F
'   �쐬�ҁ@�@�@�F������� �R���E�G���W�j�A�����O
'   �쐬�N�����@�F�Q�O�O�P�^�P�Q�^�P�S
'   �X�V�����@�@�F
'
'******************************************************************

'�ځ@�I�@�@�F�P���R�[�h�̂b�r�u�f�[�^���t�B�[���h�P�ʂɕϊ�����B
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F(i)���R�[�h�f�[�^(String)
'�@�@�@�@�@�@(o)���o�����t�B�[���h�f�[�^�i�i�[�͔z��j
'�@�@�@�@�@�@(o)���o�����t�B�[���h��
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O ����
'�쐬�N�����F�Q�O�O�P�^�P�Q�^�P�S
'�X�V�����@�F
'
Public Sub S_CSVtoTEXT(Recdata, FieldData() As Variant, FldCnt)

    Dim End_flg
    Dim fld_Data As String
    Dim Data_wk As Variant
    Dim RecLen
    
    Dim move_cnt    '�ړ��ʒu
    Dim next_cnt    '���ʒu
    Dim entcnt      '���݈ʒu
    
    Dim DQcnt As Integer '�_�u���N�H�[�e�[�V������
    
    On Error Resume Next
    
    ReDim Preserve FieldData(0)
   
    End_flg = False
    FldCnt = 0
    entcnt = 0
    RecLen = Len(Recdata)
    Do While entcnt <= RecLen And (End_flg = False)
        If InStr(entcnt + 1, Recdata, Chr(34)) = (entcnt + 1) Then
            '�擪����="
            FldCnt = FldCnt + 1
            ReDim Preserve FieldData(FldCnt)
            move_cnt = entcnt + 2
            Do While move_cnt <= RecLen
                '"��,����
                next_cnt = InStr(move_cnt, Recdata, (Chr(34) & Chr(44)))
                If next_cnt = 0 Then
                    '�Ȃ�
                    If InStr(RecLen, Recdata, Chr(34)) = RecLen Then
                        '�㕶��="
                        Data_wk = Mid(Recdata, entcnt + 2, RecLen - entcnt - 2)
                        S_CountDQ Data_wk, DQcnt    '"���Z�o
                        If (DQcnt Mod 2) = 0 Then
                            '����
                            S_CmbData Data_wk, fld_Data  '�f�[�^�ϊ�
                            Data_wk = fld_Data
                        Else
                            Data_wk = Mid(Recdata, entcnt + 1)   '�c�f�[�^���o
                        End If
                    Else
                        '�㕶��<>"
                        Data_wk = Mid(Recdata, entcnt + 1)   '�c�f�[�^���o
                    End If
                    FieldData(FldCnt) = Data_wk
                    End_flg = True
                    Exit Do
                Else
                    '����
                    Data_wk = Mid(Recdata, entcnt + 2, next_cnt - entcnt - 2)
                    S_CountDQ Data_wk, DQcnt    '"���Z�o
                    If (DQcnt Mod 2) = 0 Then
                        '����
                        S_CmbData Data_wk, fld_Data  '�f�[�^�ϊ�
                        FieldData(FldCnt) = fld_Data
                        entcnt = next_cnt + 1
                        Exit Do
                    Else
                        '�
                        move_cnt = next_cnt + 2
                    End If
                End If
            Loop
        Else
            '�擪����<>"
            FldCnt = FldCnt + 1
            ReDim Preserve FieldData(FldCnt)
            If InStr(entcnt + 1, Recdata, Chr(44)) = (entcnt + 1) Then
                '�擪����=,
                FieldData(FldCnt) = ""
                entcnt = entcnt + 1
            Else
                '�擪����<>","
                next_cnt = InStr(entcnt + 2, Recdata, Chr(44))
                '���̃J���}����H
                If next_cnt = 0 Then
                    '�Ȃ��^�c��S������
                    Data_wk = Mid(Recdata, entcnt + 1)
                    End_flg = True
                Else
                    '����^�I��
                    Data_wk = Mid(Recdata, entcnt + 1, next_cnt - entcnt - 1)
                    entcnt = next_cnt
                End If
                FieldData(FldCnt) = Data_wk
            End If
        End If
    Loop
    
End Sub

'�ځ@�I�@�@�F�_�u���N�H�[�e�[�V�����̐��𐔂���B
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F(i)���̓f�[�^
'�@�@�@�@�@�@(o)�_�u���N�H�[�e�[�V�����̐�
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O ����
'�쐬�N�����F�Q�O�O�P�^�P�Q�^�P�S
'�X�V�����@�F
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
            '�Ȃ�
            Exit Do
        End If
        Pos = Pos_next
    Loop
       
End Sub

'�ځ@�I�@�@�F�t�B�[���h�f�[�^�̕ϊ�
'���@���@�@�F
'���@�ʁ@�@�F
'���@���@�@�F(i)�t�B�[���h�f�[�^
'�@�@�@�@�@�@(o)�ϊ���̃t�B�[���h�f�[�^
'�߂�l�@�@�F
'�쐬�ҁ@�@�F������� �R���E�G���W�j�A�����O ����
'�쐬�N�����F�Q�O�O�P�^�P�Q�^�P�S
'�X�V�����@�F
'
Private Sub S_CmbData(Data As Variant, CmbData As String)
   
    Dim Pos As Integer
    Dim wData As String
    Dim Pos_next  As Integer
   
    On Error Resume Next
   
    '""->"�u��
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

