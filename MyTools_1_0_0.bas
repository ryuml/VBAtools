Attribute VB_Name = "MyTools_1_0_0"
Option Explicit


Sub ActivateFirstSheet_WindowZoom100_SelectAnyCell()
'
' ActivateFirstSheet_WindowZoom100_SelectAnyCell Macro
' �u�b�N�̊e�V�[�g�̊g�嗦��100%�ɓ��ꂵ�A�I���Z����A1�A�V�[�g�͐擪�V�[�g��I��������Ԃɂ��܂��B
'

'
    Dim zoom_val As Integer, selected_cell As String
    zoom_val = 100
    selected_cell = "A1"
    
    Dim i As Long
    For i = 0 To Worksheets.Count - 1
        Worksheets(Worksheets.Count - i).Activate
        ActiveWindow.Zoom = zoom_val
        Range(selected_cell).Select
    Next i
End Sub



Sub Set_FontAndCellSize_inSheet_MeiryoUI()
'
' Set_FontAndCellSize_inSheet_MeiryoUI Macro
' �V�[�g���̂��ׂẴZ���̃t�H���g��A�`D��̃Z���̃T�C�Y��ݒ肵�܂��B
'

'
    '�A�N�e�B�u�V�[�g��̂��ׂẴt�H���g��ݒ�
    Cells.Select
    With Selection.Font
        .Name = "Meiryo UI"
        '.Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    '�A�N�e�B�u�V�[�g��̂��ׂĂ̗�̕���ݒ�
    Cells.Select
    Dim col_width_num As Double  '��̕��̒l
    col_width_num = 8.1    'Meiryo UI �̎��̃f�t�H���g�̗�̕��̒l
    Selection.ColumnWidth = col_width_num
    
    'A��`D��̕���ݒ�
    'Meiryo UI �̍s�̍����̃f�t�H���g�l(�|�C���g�l) = 15�|�C���g = 25�s�N�Z�� = 1.8(��̕��̒l)
    ' �� A��`D��̃Z�������ׂĐ����`�ɂ��邽�߂̐ݒ�
    col_width_num = 1.8
    Columns("A:D").Select
    Selection.ColumnWidth = col_width_num
    
    '�A�N�e�B�u�V�[�g��̂��ׂĂ̍s�̍�����ݒ�
    Cells.Select
    Dim point_num As Double  '�|�C���g�l(�s�̍����̒l)
    '��̕��̒l(�W���t�H���g�̕�����Ƃ������ȒP��)���|�C���g(VBA��Excel�����̊�{�P��)�֕ϊ�
    point_num = Range("A1").Width
    Selection.RowHeight = point_num
    
    '�I������
    Range("A1").Select
    
End Sub



Sub Copy_Sheet_x5()
'
' Copy_Sheet_x5 Macro
' ���݂̃V�[�g�̃R�s�[��5�V�[�g���������܂��B
'

'
    Dim sheet_name As String, sheet_num As Integer
    sheet_name = ActiveSheet.Name
    sheet_num = Sheets(sheet_name).Index
    
    '
    'Worksheets (sheet_num)
    '
    
    Dim i As Long
    For i = 1 To 5
        Sheets(sheet_name).Select
        Sheets(sheet_name).Copy After:=Sheets(i)
    Next i
    
End Sub



Sub Copy_Sheet_x10()
'
' Copy_Sheet_x10 Macro
' ���݂̃V�[�g�̃R�s�[��10�V�[�g���������܂��B
'

'
    Dim sheet_name As String, sheet_num
    sheet_name = ActiveSheet.Name
    sheet_num = Sheets(sheet_name).Index
    
    '
    'Worksheets (sheet_num)
    '
    
    Dim i As Long
    For i = 1 To 10
        Sheets(sheet_name).Select
        Sheets(sheet_name).Copy After:=Sheets(i)
    Next i
    
End Sub



Function ExistsSheet(sheetName As String) As Boolean
    Dim sh  As Worksheet

    '�߂�l�̏����l�FFalse
    ExistsSheet = False

    On Error Resume Next

    Set sh = Worksheets(sheetName)

    '�G���[�l���Ȃ��ꍇ�F���[�N�V�[�g����
    If Err.Number = 0 Then ExistsSheet = True

End Function



Sub GetLinksDirAllFiles()
' �w��̃f�B���N�g������Office�t�@�C�����̃����N���擾

    Dim WriteWS         As Worksheet
    Dim FolderPath      As String
    Dim FilePath        As String
    Dim FileName        As String
    Dim ExcelExt1       As String
    Dim ExcelExt2       As String
    Dim RowCnt          As Integer
    Dim WriteWS_Name    As String
    
    ' Excel�̊g���q�w��
    ExcelExt1 = ".xls"
    ExcelExt2 = ".xlsx"
    ' �������݃��[�N�V�[�g���ݒ�
    WriteWS_Name = "LinkList"
    
    If ExistsSheet(WriteWS_Name) Then
        Worksheets(WriteWS_Name).Activate
        ActiveSheet.Cells.Clear
    Else
        '// �u�b�N�̈�ԍ��ɐV�K�V�[�g��ǉ�
        Call Worksheets.Add(Before:=Worksheets(1))
        'Worksheets(1).Activate
        ActiveSheet.Name = WriteWS_Name   ' ���O�ύX
    End If
    
    '�}�N���t�@�C���̃V�[�g��ݒ�
    Set WriteWS = ThisWorkbook.ActiveSheet
    
    ' �s�J�E���g������
    RowCnt = 1
    ' �񖼐ݒ�
    WriteWS.Cells(RowCnt, 1).Activate
    '// A��F�u�b�N��
    ActiveCell.Value = "�u�b�N��"
    ActiveCell.ColumnWidth = 20
    ActiveCell.Font.Bold = True
    '// B��F�V�[�g��
    ActiveCell.Offset(0, 1).Value = "�V�[�g��"
    ActiveCell.Offset(0, 1).ColumnWidth = 20
    ActiveCell.Offset(0, 1).Font.Bold = True
    '// C��F���W
    ActiveCell.Offset(0, 2).Value = "���W"
    ActiveCell.Offset(0, 2).ColumnWidth = 10
    ActiveCell.Offset(0, 2).Font.Bold = True
    '// D��F��ނ��o��
    ActiveCell.Offset(0, 3).Value = "���"
    ActiveCell.Offset(0, 3).ColumnWidth = 10
    ActiveCell.Offset(0, 3).Font.Bold = True
    '// E��F�A�h���X
    ActiveCell.Offset(0, 4).Value = "�A�h���X"
    ActiveCell.Offset(0, 4).ColumnWidth = 20
    ActiveCell.Offset(0, 4).Font.Bold = True
    ' �s�J�E���g�X�V
    RowCnt = RowCnt + 1
    
    '�t�H���_�I��p�_�C�A���O��\��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path '���݂̃t�H���_�p�X
        If .Show = False Then Exit Sub
        FolderPath = .SelectedItems(1) '�t�H���_�p�X���擾
    End With
    
    FileName = Dir(FolderPath & "\*" & ExcelExt1) '�ŏ��̃t�@�C�������擾
    
    Do While FileName <> ""
        If (LCase(FileName) Like ("*" & ExcelExt1) _
            Or LCase(FileName) Like ("*" & ExcelExt2)) Then
            '�t�@�C�����J��
            Workbooks.Open FileName:=FolderPath & "\" & FileName
            
            '�J�����t�@�C���̃����N�𒲂ׂāA�������݃V�[�g�֋L�^
            'WriteWS.Cells(i, 1) = Workbooks(FileName).ActiveSheet.Cells(1, 1)
            Call ExcelExtractLinks(WriteWS, Workbooks(FileName), RowCnt)
            '�J�����t�@�C�������
            Application.DisplayAlerts = False
            Workbooks(FileName).Close
            Application.DisplayAlerts = True
        End If
        
        ' �X�V����
        FileName = Dir() '���̃t�@�C�������擾
    Loop
    
End Sub



Sub ExcelExtractLinks(ByRef w_ws As Worksheet, ByRef r_wb As Workbook, _
                        ByRef row_cnt As Integer)
' �A�N�e�B�u�u�b�N�̑S�V�[�g�̃Z���Ɖ摜�̃n�C�p�[�����N�𔲂��o���A�V�K�V�[�g�Ɉꗗ�ŏo��
    
    Dim sht             As Worksheet    '// ���[�N�V�[�g
    Dim ar()            As String       '// �n�C�p�[�����N�z��
    Dim hLink           As Hyperlink    '// �n�C�p�[�����N
    Dim sCellAddress    As String       '// �Z�����W
    Dim sLinkAddress    As String       '// �����N��
    Dim sType           As String       '// ���
    Dim s               As Variant      '// �z��̗v�f������
    Dim v               As Variant      '// ����
    Dim TmpUR           As Range        '// �ꎟ�I��UsedRange�̃f�[�^�͈�
    Dim CellDataRange   As Range        '// �Z���̃f�[�^�͈͕���
    Dim CellData        As Range        '// �Z���̃f�[�^�͈͒P��
    
    ReDim ar(0)
    ReDim CellDataArr(0)
    
    '// �A�N�e�B�u�u�b�N�̑S�V�[�g�����[�v
    For Each sht In r_wb.Worksheets
        '// �V�[�g���̃n�C�p�[�����N�����[�v
        For Each hLink In sht.Hyperlinks
            '// Range�i�Z���j�̏ꍇ
            If hLink.Type = msoHyperlinkRange Then
                sCellAddress = hLink.Range.Address(False, False)
                sType = "Range"
            '// Shape�i�摜�j�̏ꍇ
            ElseIf hLink.Type = msoHyperlinkShape Then
                sCellAddress = hLink.Shape.TopLeftCell.Address(False, False)
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
            
            '// �u�b�N���{�V�[�g���{�Z�����W�{��ށ{�A�h���X
            ar(UBound(ar)) = r_wb.Name & vbTab & sht.Name & vbTab & sCellAddress _
                                & vbTab & sType & vbTab & sLinkAddress
            ReDim Preserve ar(UBound(ar) + 1)
        Next
        
On Error GoTo SkipNoFormula
        '// �͈͎擾
        '// �������珑���E�f�[�^�������Ă���͈͂��擾
        Set TmpUR = sht.UsedRange
        If Not TmpUR Is Nothing Then
            If IsEmpty(TmpUR) Then
                '// �Z�����e���󔒁i�f�[�^���� or �󕶎���j���������ꍇ�̓G���[����̂��ߐ������肹���A���������Ȃ��ݒ�
                Set CellDataRange = Nothing
            Else
                '// CellDataRange �ɔ͈́i�I�u�W�F�N�g�Q�Ɓj������ꍇ�A����������Z���͈͂��擾
                Set CellDataRange = TmpUR.SpecialCells(xlCellTypeFormulas)
            End If
        End If
        
        '// HYPERLINK�̒l�擾
        If Not CellDataRange Is Nothing Then
            '// CellDataRange �ɔ͈́i�I�u�W�F�N�g�Q�Ɓj������ꍇ
            For Each CellData In CellDataRange
                If CellData.Formula Like "*HYPERLINK*" Then
                    '// "HYPERLINK" ���܂܂�Ă���ꍇ
                    sCellAddress = CellData.Address(False, False)
                    sType = "Formula"
                    sLinkAddress = CellData.Formula
                    sLinkAddress = Replace(sLinkAddress, "=", "", Count:=1)
                    
                    '// �u�b�N���{�V�[�g���{�Z�����W�{��ށ{�A�h���X
                    ar(UBound(ar)) _
                        = r_wb.Name & vbTab & sht.Name & vbTab & sCellAddress _
                            & vbTab & sType & vbTab & sLinkAddress
                    ReDim Preserve ar(UBound(ar) + 1)
                End If
            Next
        End If
SkipNoFormula:
        '// �����Z�������������ꍇ�X�L�b�v
        Err.Clear
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