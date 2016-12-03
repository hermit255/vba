Option Explicit
  Dim t As table '�e�[�u���I�u�W�F�N�g
  Dim region As Range

  Sub insert_general_forms()
    Set t = New table
    Set region = Cells(3, 1).CurrentRegion '�e�[�u���͈͂̒�`
    region.Select
    
    setPreferance_forms
    Call t.readRangeH(region)
    renderSQL (0)
    Call t.report

  End Sub
  
  Sub insert_master_forms()
    Set t = New table
    setPreferance_forms
    '�e�[�u���͈͂��Œ�
    Set region = Range(Cells(3.1), Cells(43, 2))
    Call t.readRangeV(region)    '�J��������������
    renderSQL (0)
    Call t.report

  End Sub
  Sub delete_general_forms()
    Set t = New table
    Set region = Cells(3, 1).CurrentRegion '�e�[�u���͈͂̒�`
    
    setPreferance_forms
    Call t.readRangeH(region)    '�J��������������
    renderSQL (1)

  End Sub
  Sub insertSelectedH()
    Dim name As String
    Dim db As String
    
    Set t = New table
    
    'DB����`
    db = "dbname"
    Call t.setDatabase(db)
    
    'B1�Z���Ƀe�[�u�������L��������̂Ƃ���
    name = "tableName"
    Call t.setName(name)
    
    '�e�[�u���͈͂��Œ�
    Set region = Selected
    Call t.readRangeV(region)    '�J��������������
    renderSQL (0)
    Call t.report
  End Sub
  Private Sub setPreferance_forms()
    Dim name As String
    Dim db As String
    
    'DB����`
    db = "serious"
    Call t.setDatabase(db)
    
    'B1�Z���Ƀe�[�u�������L��������̂Ƃ���
    name = ActiveSheet.Cells(1, 2).value
    Call t.setName(name)

  End Sub
  Private Sub renderSQL(Optional mode As Integer = 0)
    Dim tCell As Range 'SQL���̏������ݐ� targetCell
  
    '�e�[�u�������̉��ɁA��������SQL���o��
    Set tCell = Cells(65535, 1).End(xlUp).Offset(2, 0)
    'mode �ɂ��SQL��ނ𕪂���
    If (mode = 1) Then
      tCell.value = t.delete(, 1)
    ElseIf (mode = 2) Then
      tCell.value = t.update
    Else
      tCell.value = t.insert
    End If
    tCell.Select
  End Sub
