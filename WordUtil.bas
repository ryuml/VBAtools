Attribute VB_Name = "WordUtil"
Option Explicit


' �w���Word�t�@�C�������ɊJ����Ă��邩�m�F
Function CheckIfWordFileIsOpen(ByVal FilePath As String) As Boolean
    ' �֘A����I�u�W�F�N�g�̐錾
    Dim wordApp As Object
    Dim doc As Object
    Dim IsDocOpen As Boolean
    
    IsDocOpen = False   ' ������
    
    ' �G���[���������Ă��X�N���v�g�̎��s�𑱍s
    On Error Resume Next
    ' ������Word�A�v���P�[�V���������邩�m�F
    Set wordApp = GetObject(, "Word.Application")
    ' �G���[�n���h�����O�����ɖ߂�
    On Error GoTo 0
    
    ' Word�A�v���P�[�V���������݂���ꍇ
    If Not wordApp Is Nothing Then
        ' �J����Ă���S�Ẵh�L�������g�ɑ΂��ď��������s
        For Each doc In wordApp.Documents
            If doc.Path & "\" & doc.Name = FilePath Then
                ' �t�@�C�����J����Ă���ꍇ��True��Ԃ��A�֐����I��
                IsDocOpen = True
                ' ���ɊJ����Ă���ꍇ�͏����𒆎~
                Exit For
                End
            End If
        Next doc
    Else
        ' Word�A�v���P�[�V���������s����Ă��Ȃ��ꍇ
    End If
    
    ' �I�u�W�F�N�g�̉��
    Set doc = Nothing
    Set wordApp = Nothing
    
    ' ���ʕԋp
    CheckIfWordFileIsOpen = IsDocOpen
End Function
