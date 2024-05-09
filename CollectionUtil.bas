Attribute VB_Name = "CollectionUtil"
Option Explicit

'===========================================================
'
' �R���N�V��������p���W���[��
'
' [�����T�v]
'�@�E�R���N�V�����̃L�[�����A�A�C�e������
'
' [����]
'  �� 1. ExistsKey�iCollention���̃L�[�����֐��j
'      �E��Q�������L�[�Ƃ���Item���\�b�h�����s���A
' �@�@ �@���ʂ����ƂɃL�[�̑��݂��m�F����B
'  �� 2. ExistsItem�iCollention�̊i�[�f�[�^�̑��݃`�F�b�N�֐��j
'      �E��Q�����������ΏۂƂ��Ċe�����o�[�Ɠˍ����A
' �@�@ �@�����o�[���̑��݃`�F�b�N���ʂ�Ԃ��B
'  �� 3. ExistsNoKeyItem�iKey�����w��ł�ExistsItem�֐��j
'      �E��Q�����������ΏۂƂ��đ�������s���A
' �@�@ �@�G���[���ʂ���ɁA���݃`�F�b�N���ʂ�Ԃ��B
'
'===========================================================

' ���W���[����
Const MODULE_NAME = "CollectionUtil"




'*********************************************************
'* ExistsKey�iCollention���̃L�[�����֐��j
'*********************************************************
'* ��P���� | Collection | �����ΏۂƂȂ�I�u�W�F�N�g
'* ��Q���� |   String   | ��������L�[
'*  �߂�l�@|   Boolan   | True Or False ��False�������l
'*********************************************************
'*   ����   | ��Q�������L�[�Ƃ���Item���\�b�h�����s���A
'*   �@�@   | ���ʂ����ƂɃL�[�̑��݂��m�F����B
'*********************************************************
'*   ���l   | �I�u�W�F�N�g���ݒ�̏ꍇ �� �߂�l�uFalse�v
'*   �@�@   | �����o�[���u0�v�̏ꍇ �� �߂�l�uFalse�v
'*********************************************************
 
Function ExistsKey(objCol As Collection, strKey As String) As Boolean
     
    '�߂�l�̏����l�FFalse
    ExistsKey = False
     
    '�ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    'Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
     
    On Error Resume Next
     
    'Item���\�b�h�����s
    Call objCol.Item(strKey)
         
    '�G���[�l���Ȃ��ꍇ�F�L�[�����̓q�b�g�i�߂�l�FTrue�j
    If Err.Number = 0 Then ExistsKey = True
 
End Function



'*********************************************************
'* ExistsItem�iCollention�̊i�[�f�[�^�̑��݃`�F�b�N�֐��j
'*********************************************************
'* ��P���� | Collection | �����ΏۂƂȂ�I�u�W�F�N�g
'* ��Q���� |  Variant   | ��������f�[�^
'*�@�߂�l  |   Boolan   | True Or False ��False�������l
'*********************************************************
'*   ����   | ��Q�����������ΏۂƂ��Ċe�����o�[�Ɠˍ����A
'*   �@�@   | �����o�[���̑��݃`�F�b�N���ʂ�Ԃ��B
'*********************************************************
'*   ���l   | �I�u�W�F�N�g���ݒ�̏ꍇ �� �߂�l�uFalse�v
'*   �@�@   | �����o�[���u0�v�̏ꍇ �� �߂�l�uFalse�v
'*********************************************************
 
Function ExistsItem(objCol As Collection, varItem As Variant) As Boolean
     
    Dim v As Variant
     
    '�߂�l�̏����l�FFalse
    ExistsItem = False
     
    '�ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    'Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
     
    'Collection�̊e�����o�[�Ɠˍ�
    For Each v In objCol
         
        '�ˍ����ʂ���v�����ꍇ�F�߂�l�uTrue�v�Ƀ��[�v����
        If v = varItem Then ExistsItem = True: Exit For
         
    Next
     
End Function



'*********************************************************
'* ExistsNoKeyItem�iKey�����w��(�C���f�b�N�X�`��)��
'  Collention�̊i�[�f�[�^�̑��݃`�F�b�N�֐��j
'*********************************************************
'* ��P���� | Collection | �����ΏۂƂȂ�I�u�W�F�N�g
'* ��Q���� |  Variant   | ��������f�[�^
'*�@�߂�l  |   Boolan   | True Or False ��False�������l
'*********************************************************
'*   ����   | ��Q�����������ΏۂƂ��đ�������s���A
'*   �@�@   | �G���[���ʂ���ɁA���݃`�F�b�N���ʂ�Ԃ��B
'*********************************************************
'*   ���l   | �I�u�W�F�N�g���ݒ�̏ꍇ �� �߂�l�uFalse�v
'*   �@�@   | �����o�[���u0�v�̏ꍇ �� �߂�l�uFalse�v
'*********************************************************
 
Function ExistsNoKeyItem(objCol As Collection, varItem As Variant) As Boolean
     
    Dim v As Variant
     
    '�߂�l�̏����l�FFalse
    ExistsNoKeyItem = False
     
    '�ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    'Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
    
    On Error Resume Next
    
    Set v = objCol(varItem)
    
    '�G���[�l���Ȃ��ꍇ�F���[�N�V�[�g����
    If Err.Number = 0 Then ExistsNoKeyItem = True
    
End Function
