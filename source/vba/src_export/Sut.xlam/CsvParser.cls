VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' CSV��̓N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/03�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' ��؂蕶��
Private m_separator As String

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ������������
' =========================================================
Public Sub init(Optional ByVal separator As String = ",")
    m_separator = separator
End Sub

' =========================================================
' ��CSV��͏���
' =========================================================
Public Function parse(ByVal csvStr As String) As ValCollection

    Dim ret    As New ValCollection
    Dim retRec As New ValCollection

    Dim i As Long
    
    Dim outNextPos As Long
    Dim outFieldStr As String
    Dim outIsNextRec As Boolean
    
    Dim csvChar   As String
    Dim csvStrLen As Long
    csvStrLen = Len(csvStr)

    i = 1
    Do While i <= csvStrLen
    
        ' �t�B�[���h����͂���
        parseOneField csvStr, i, outNextPos, outFieldStr, outIsNextRec
        
        ' ���R�[�h�Ƀt�B�[���h��ǉ�����
        retRec.setItem outFieldStr
        
        ' �C���f�b�N�X��i�߂�
        i = outNextPos
        
        If outIsNextRec Then
        
            ' ���R�[�h�̏I�[�ɒB���āA���I�[����؂蕶���̏ꍇ�́A�����ɋ󕶎����ǉ�����
            If (Mid$(csvStr, i - 3, 1) = m_separator And (Mid$(csvStr, i - 2, 1) = vbCr And Mid$(csvStr, i - 1, 1) = vbLf)) Or _
               (Mid$(csvStr, i - 2, 1) = m_separator And (Mid$(csvStr, i - 1, 1) = vbCr Or Mid$(csvStr, i - 1, 1) = vbLf)) Then
            
                ' ���R�[�h�Ƀt�B�[���h��ǉ�����
                retRec.setItem ""
                
            End If
            
            ' ���̍s�Ɉړ�����̂Ń��R�[�h���X�g�Ƀ��R�[�h��ǉ�����
            ret.setItem retRec
            ' ���R�[�h�𐶐�����
            Set retRec = New ValCollection
            
        Else
        
            If i > csvStrLen Then
                ' ���R�[�h�̏I�[�ɒB���āA���I�[����؂蕶���̏ꍇ�́A�����ɋ󕶎����ǉ�����
                If Mid$(csvStr, csvStrLen, 1) = m_separator Then
                
                    ' ���R�[�h�Ƀt�B�[���h��ǉ�����
                    retRec.setItem ""
                    
                End If
            
            End If
            
        End If
        
    Loop
    
    If retRec.count > 0 Then
        ' ���̍s�Ɉړ�����̂Ń��R�[�h���X�g�Ƀ��R�[�h��ǉ�����
        ret.setItem retRec
    End If
    
    Set parse = ret

End Function

' =========================================================
' ��CSV�t�B�[���h��͏���
' =========================================================
Private Sub parseOneField(ByVal csvStr As String _
                        , ByVal first As Long _
                        , ByRef outNextPos As Long _
                        , ByRef outFieldStr As String _
                        , ByRef outIsNextRec As Boolean)

    Dim isEnclosed As Boolean

    Dim i As Long
    
    Dim csvCharBefore   As String
    Dim csvCharBefore2  As String
    Dim csvCharNext     As String
    Dim csvChar         As String
    Dim csvStrLen As Long
    csvStrLen = Len(csvStr)
    
    outNextPos = first
    outFieldStr = ""
    outIsNextRec = False


    For i = first To csvStrLen
    
        csvChar = Mid$(csvStr, i, 1)
        
        If i = first Then
            ' �ŏ��̕���
        
            If csvChar = """" Then
                ' �ŏ��̕������͂ݕ����ł��邩�𔻒肷��
                isEnclosed = True

            ElseIf csvChar = m_separator Then
                ' ���݂̕�������؂蕶���̏ꍇ
                ' �t�B�[���h��͏I��
                outFieldStr = Mid$(csvStr, first, i - first)
                
                isNextRec csvStr, i + 1, outNextPos, outIsNextRec
                
                Exit Sub

            Else
            
                isNextRec csvStr, i, outNextPos, outIsNextRec
                If outIsNextRec Then
                    Exit Sub
                End If
            
            End If
            
        Else
            ' 2�����ڈȍ~
        
            If isEnclosed Then
                ' �͂ݕ����̏ꍇ
                
                csvCharNext = Mid$(csvStr, i + 1, 1)
                
                If ( _
                    ( _
                      csvCharBefore <> """" Or _
                     (csvCharBefore = """" And csvCharBefore2 = """") Or _
                     (csvCharBefore = """" And csvCharBefore2 = "") _
                    )) And _
                     csvChar = """" And _
                    (csvCharNext = m_separator Or csvCharNext = vbCr Or csvCharNext = vbLf Or csvCharNext = "") Then
                    ' �ȉ��̏����𖞂�������
                    ' �E���݂̕������͂ݕ����ł���
                    ' �E���O�̕������͂ݕ����ł͂Ȃ� �܂��� �͂ݕ����ł����Ă�2�����A���ň͂ݕ����ł���
                    ' �E����̕�������؂蕶���ŏI���A���s�R�[�h�ŏI���A�����̏I�[�ł���
                    
                    ' �t�B�[���h��͏I��
                    
                    outNextPos = i + 1 ' ���񕶎��A��؂蕶���Ƒ����͂��Ȃ̂ŁA1������΂�
                    outFieldStr = replace(Mid$(csvStr, first + 1, i - 1 - first), """""", """")
                    
                    isNextRec csvStr, i + 2, outNextPos, outIsNextRec

                    Exit Sub
                
                End If
            Else
                ' �͂ݕ����Ȃ�

                If csvChar = m_separator Then
                    ' ���݂̕�������؂蕶���̏ꍇ
                    ' �t�B�[���h��͏I��
                    outFieldStr = Mid$(csvStr, first, i - first)
                    isNextRec csvStr, i + 1, outNextPos, outIsNextRec
                    
                    Exit Sub
                    
                Else
                
                    isNextRec csvStr, i, outNextPos, outIsNextRec
                    If outIsNextRec Then
                        outFieldStr = Mid$(csvStr, first, i - first)
                        Exit Sub
                    End If
                
                End If
                
            
            End If
            
        End If
        
        csvCharBefore2 = csvCharBefore
        csvCharBefore = csvChar
        
    Next
    
    outNextPos = i
    outFieldStr = Mid$(csvStr, first, i - 1)
    outIsNextRec = True

End Sub

Private Sub isNextRec(ByVal csvStr As String _
                        , ByVal first As Long _
                        , ByRef outNextPos As Long _
                        , ByRef outIsNextRec As Boolean)

    outIsNextRec = False
    outNextPos = first

    Dim csvCharNext     As String
    Dim csvChar         As String
    
    csvChar = Mid$(csvStr, first, 1)
    csvCharNext = Mid$(csvStr, first + 1, 1)

    If csvChar = vbCr And csvCharNext = vbLf Then
        ' CRLF�ŏI������ꍇ
        outIsNextRec = True
        outNextPos = first + 2
    ElseIf csvChar = vbCr Or _
           csvChar = vbLf Then
        ' CR or LF�ŏI������ꍇ
        outIsNextRec = True
        outNextPos = first + 1
    End If

End Sub

Private Function excludeEncloseChar(ByVal field As String) As String

    If InStr(field, """") = 1 And InStrRev(field, """") = Len(field) Then
        ' �͂ݕ����ň͂܂�Ă���ꍇ�͎�菜��
        excludeEncloseChar = Mid$(field, 1, Len(field) - 2)
    Else
        ' �͂ݕ����̏ꍇ�͂��̂܂�
        excludeEncloseChar = field
    End If


End Function

