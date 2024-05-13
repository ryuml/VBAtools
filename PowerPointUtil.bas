Attribute VB_Name = "PowerPointUtil"
Option Explicit


' �w���PowerPoint�t�@�C�������ɊJ����Ă��邩�m�F
Function CheckIfPwPtFileIsOpen(ByVal FilePath As String) As Boolean
    ' �֘A����I�u�W�F�N�g�̐錾
    Dim ppApp As Object
    Dim pres As Object   ' Presentation
    Dim IsPresOpen As Boolean
    
    IsPresOpen = False   ' ������
    
    ' �G���[���������Ă��X�N���v�g�̎��s�𑱍s
    On Error Resume Next
    ' ������PowerPoint�A�v���P�[�V���������邩�m�F
    Set ppApp = GetObject(, "PowerPoint.Application")
    ' �G���[�n���h�����O�����ɖ߂�
    On Error GoTo 0
    
    ' PowerPoint�A�v���P�[�V���������݂���ꍇ
    If Not ppApp Is Nothing Then
        ' �J����Ă���S�Ẵv���[���e�[�V�����ɑ΂��ď��������s
        For Each pres In ppApp.Presentations
            If pres.Path & "\" & pres.Name = FilePath Then
                ' �t�@�C�����J����Ă���ꍇ��True��Ԃ��A�֐����I��
                IsPresOpen = True
                ' ���ɊJ����Ă���ꍇ�͏����𒆎~
                Exit For
                End
            End If
        Next pres
    Else
        ' PowerPoint�A�v���P�[�V���������s����Ă��Ȃ��ꍇ
    End If
    
    ' �I�u�W�F�N�g�̉��
    Set pres = Nothing
    Set ppApp = Nothing
    
    ' ���ʕԋp
    CheckIfPwPtFileIsOpen = IsPresOpen
End Function



Sub GetPpAppHlinks(ByRef w_ws As Worksheet, ByRef objPres As Object, ByRef row_cnt As Integer)

' �A�N�e�B�u�u�b�N�̑S�V�[�g�̃Z���Ɖ摜�̃n�C�p�[�����N�𔲂��o���A�V�K�V�[�g�Ɉꗗ�ŏo��
    
    Dim ar()            As String       '// �n�C�p�[�����N�z��
    Dim hLink           As Hyperlink    '// �n�C�p�[�����N(PowerPoint.Hyperlink)
'    Dim sCellAddress    As String       '// �Z�����W
    Dim sLinkAddress    As String       '// �����N��
    Dim sType           As String       '// ���
    Dim s               As Variant      '// �z��̗v�f������
    Dim v               As Variant      '// ����
'    Dim ppApp       As Object   'PowerPoint.Application
'    Dim p As Integer, y As Integer   'p:�X���C�h�y�[�W Excel:y�s
    Dim SldPage     As Integer
    Dim objShape    As Object 'PowerPoint.Shape '�p���|�̃V�F�C�v�A�e�L�X�g�A�}�`�ق�
    Dim objSlide    As Object 'PowerPoint.Slide

'    On Error Resume Next  '�擾�G���[���Ɏ���
'    Set ppApp = GetObject(, "PowerPoint.Application")
'    On Error GoTo 0  '�G���[�����ɖ߂��������Y���ƁA�f�o�b�O���Ƀn�}�邩�璍��

    If ppApp Is Nothing Then
        MsgBox "�p���[�|�C���g���擾�ł��܂���B"
        Exit Sub
    End If
    
    ReDim ar(0)

    '// �A�N�e�B�u�v���[���e�[�V�����̑S�X���C�h�����[�v
    For Each objSlide In objPres.Slides
        '// �X���C�h���̃n�C�p�[�����N�����[�v
        For Each hLink In objSlide.Hyperlinks
            '// Range�i�Z���j�̏ꍇ
            If hLink.Type = msoHyperlinkRange Then
                sType = "Range"
            '// Shape�i�摜�j�̏ꍇ
            ElseIf hLink.Type = msoHyperlinkShape Then
                sType = "Shape"
            End If
            
            '// �O�������N���ݒ肳��Ă���ꍇ
            If hLink.Address <> "" Then
                '// �O���ւ̃n�C�p�[�����N���擾
                sLinkAddress = hLink.Address
            '// ���������N���ݒ肳��Ă���ꍇ
            Else
                '// �����ւ̃n�C�p�[�����N���擾
                sLinkAddress = hLink.SubAddress
            End If
            
            '// �t�@�C�����{�X���C�h���{�X���C�h�ԍ��{��ށ{�A�h���X
            ar(UBound(ar)) = objPres.Name & vbTab & hLink.TextToDisplay & vbTab & objSlide.SlideNumber _
                                & vbTab & sType & vbTab & sLinkAddress
            ReDim Preserve ar(UBound(ar) + 1)
        Next
    Next
    
    '// �z��(0)�Ɋi�[�ς݂̏ꍇ(ar(0)���󕶎��ł͂Ȃ�)
    If ar(0) <> "" Then
        '// �]���ȗ̈���폜
        ReDim Preserve ar(UBound(ar) - 1)
    End If
    
    '// �u�b�N�̈�ԍ��ɐV�K�V�[�g��ǉ�
    'Call Worksheets.Add(Before:=Worksheets(1))
    
    ' �A�N�e�B�u�Z���ݒ�
    w_ws.Activate
    w_ws.Cells(row_cnt, 1).Activate
    
    '// �n�C�p�[�����N�̐����[�v
    For Each s In ar
        If s <> "" Then
            '// TAB�����ŕ���
            v = Split(s, vbTab)
            '// A��Ƀu�b�N�����o��
            ActiveCell.Value = v(0)
            '// B��ɃV�[�g�����o��
            ActiveCell.Offset(0, 1).Value = v(1)
            '// C��ɍ��W���o��
            ActiveCell.Offset(0, 2).Value = v(2)
            '// D��Ɏ�ނ��o��
            ActiveCell.Offset(0, 3).Value = v(3)
            '// E��ɃA�h���X���o��
            ActiveCell.Offset(0, 4).Value = v(4)
            
            '// ���̃Z����I��
            row_cnt = row_cnt + 1
            ActiveCell.Offset(1, 0).Select
        End If
    Next
    
End Sub
