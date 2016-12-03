Option Explicit
  Dim name As String
  Dim database As String
  'column �� string �̏W��
  Dim columns As New Collection
  'records �� collection(record) �̏W���Arecord �� string �̏W��
  Dim records As New Collection

  '�Œ���K�v�ȏ����� readRange("�����W�I�u�W�F�N�g") ��
  '�J�����{���R�[�h�ƂȂ�Z���͈�(�J����������)���w�肷�邾���B
  '�C�ӂ� setName("�e�[�u����"),setDatabase("DB��") �Œǉ��ݒ肷�邱�Ƃ��ł���B
  '�������ł�����A�C�ӂ̃Z���� insert �� update ����SQL������������ł��p�r�̃N���X�B

  Function insert()
    Dim SQL As String 'SQL���̃x�[�X����
    Dim SQL2 As String 'SQL���̃��R�[�h����
    Dim count As Integer
    Dim record As Collection
    count = 0
    
    SQL = "INSERT INTO " & getTarget & " " & printC(columns, 2) & " VALUES "

    For Each record In records
    If (count > 0) Then SQL2 = SQL2 & ", "
      SQL2 = SQL2 & printC(record, 1)
      count = count + 1
    Next record
    
    SQL = SQL & SQL2 & ";"
    insert = SQL
    
  End Function

  Function update(Optional pKey As Integer = 1)
    Dim newSQL As String 'SQL�ɒǉ����Ă����ꎞ�p�̕�����
    Dim SQL As String '�ŏI�I�ɏo�͂���SQL��
    Dim couple As New Collection 'columns ���L�[�Arecords �� value �Ƃ���W��
    
    For Each record In records
      Set couple = New Collection
      For Each Column In columns
          couple.Add (vConvert(Column, 2) & " = " & vConvert(record(Column)))
      Next Column
      '�ŏ��̃J������Pkey�Ɖ��肵�āAUPDATE���𔭍s����
      newSQL = "UPDATE " & getTarget & " SET " & printC(couple, 0) & " WHERE " & columns(pKey) & " = " & record(columns(pKey)) & ";"
      SQL = SQL & newSQL
    Next record
    update = SQL

    If (mode = 1) Then Call report
  End Function

  Function delete(Optional pKey As Integer = 1, Optional mode = 0)
    Dim SQL As String 'SQL���̃x�[�X����
    Dim newSQL As String 'SQL�Ɍp���������V�p�[�c
    Dim record As Collection
    
    '�ŏ��̃J������Pkey�Ɖ��肵�āADELETE���𔭍s����
    For Each record In records
      newSQL = "DELETE FROM " & getTarget & " WHERE " & columns(pKey) & " = " & record(pKey) & ";"
      SQL = SQL & newSQL
    Next record
    delete = SQL
    
    If (mode = 1) Then Call report
  End Function

  Private Function printC(ByVal coll As Collection, Optional mode As Integer = 1)
    '�J�����A���R�[�h�������o���@�\ printCollection
    'mode = 1 �ŋ󔒂�Null�̑Ή���ω��ɑΉ�����
    Dim count As Integer
    Dim data As String
    Dim element As Variant
          
    data = "( "
    For Each element In coll
       '2���ڈȍ~�̓J���}�t�^
       If (count > 0) Then data = data & ", "
       If (mode = 1) Then 'mode=1 �Ȃ烌�R�[�h�l�ɃV���O���N�H�[�g�ň͂�
         data = data & vConvert(element)
       ElseIf (mode = 2) Then 'mode = 2 �Ȃ�J�����p�Ƀo�b�N�N�I�[�g�ň͂�
         data = data & vConvert(element, 1)
       Else
         data = data & element
       End If
      count = count + 1
    Next element
      data = data & " )"
      printC = data
  End Function

  Sub readRangeH(Optional tRan As Range = Null)
    Call readRange(tRan, 0)
  End Sub

  Sub readRangeV(Optional tRan As Range = Null)
    Call readRange(tRan, 1)
  End Sub

  Private Sub readRange(Optional tRan As Range = Null, Optional mode As Integer = 0)
    '�ЂƂ܂��I��͈͂̎擾���\�b�h����������A���Α��̕�������
    'mode = 1 �̏ꍇ�̓J�������c�ɕ��Ԍ`���̕\�ɑΉ�����B�i�f�t�H���g�̓J���������j
    Dim cRan As Range '�J�������ƂȂ�ŏ��̍s columnRange
    Dim rRan As Range '��ƃ��R�[�h�s recordRange
    Dim line As Integer
    
    Dim top As Integer
    Dim bottom As Integer
    Dim left As Integer
    Dim right As Integer
    
    Set columns = New Collection
    Set records = New Collection

    '�͈͎w�肪�Ȃ���ΑI��͈͂��e�[�u���Ƃ��Ĉ���
    If (IsNull(tRan)) Then Set tRan = Selection
    
    top = tRan.row
    bottom = tRan(tRan.count).row
    left = tRan.Column
    right = tRan(tRan.count).Column

    'readRangeV
    If (mode = 1) Then
      Set cRan = Range(Cells(top, left), Cells(bottom, left))
      Set columns = getCol(cRan)
      '���R�[�h�����̓ǂݎ��
      For line = (left + 1) To right
        Set rRan = Range(Cells(top, line), Cells(bottom, line))
        records.Add Item:=getCol(rRan), key:=columns(line)
      Next line
    'mode = 0 �̏ꍇ�ireadRangeH�j
    Else
      Set cRan = Range(Cells(top, left), Cells(top, right))
      Set columns = getRow(cRan)
      '���R�[�h�����̓ǂݎ��
      For line = (top + 1) To bottom
        Set rRan = Range(Cells(line, left), Cells(line, right))
        records.Add Item:=getRow(rRan)
      Next line
    End If
  End Sub
  
  Private Function getRow(ran As Range)
    Set getRow = getLine(0, ran)
  End Function

  Private Function getCol(ran As Range)
    Set getCol = getLine(1, ran)
  End Function

  Private Function getLine(mode As Integer, ran As Range)
    'mode = 1 �� getColumn 0 ��getRow �ɑΉ�
    Dim data As Collection
    Dim line As Integer
    Dim val As String
    Dim col As Integer
    Dim row As Integer
    
    Dim start As Integer
    Dim last As Integer
    
    Set data = New Collection
    If (mode = 1) Then
      start = ran.row
      last = ran(ran.count).row
      col = ran.Column
    Else
      start = ran.Column
      last = ran(ran.count).Column
      row = ran.row
    End If
    
    For line = start To last
      If (mode = 1) Then
        val = Cells(line, col).value
        If (columns.count = 0) Then
          data.Add val
        Else
          data.Add val, columns(data.count + 1)
        End If
      Else
        val = Cells(row, line).value
        If (columns.count = 0) Then
          data.Add val
        Else
          data.Add val, columns(data.count + 1)
        End If
      End If
    Next line
    Set getLine = data
  End Function

  Private Function vConvert(value As Variant, Optional mode As Integer = 0)
    '�l��SQL�p�ɕϊ� valueConvert
    Select Case value
      Case "NULL"
        value = "NULL"
      Case "Null"
        value = "NULL"
      Case "null"
        value = "NULL"
      Case "NOW()"
        value = "NOW()"
      Case ""
        value = Chr(39) & Chr(39)
      Case Else
        If (mode = 1) Then
          value = "`" & value & "`"
        Else
         value = Chr(39) & value & Chr(39)
        End If
        '�Z�����̉��s�����s�R�[�h�ɕϊ�
        value = Replace(value, vbLf, "\n")
    End Select
    vConvert = value
  End Function

  Sub setName(Optional str As String)
    name = str
  End Sub
  
  Sub setDatabase(Optional str As String)
    database = str
  End Sub
  
  Private Function getTarget()
    Dim data As String
    
    data = "`table_name`"
    If Not (IsNull(name)) Then data = "`" & name & "`"
    If Not (IsNull(database)) Then data = "`" & database & "`." & data
    getTarget = data
  End Function

  Sub report()
    '�����������̃��b�Z�[�W
    Dim msg As String
    msg = "�J���� " & columns(1) & " �` " & columns(columns.count) & " �܂ŁA�S " _
          & records.count & " ���ɂ��ď����o���܂����B"
    MsgBox msg
  End Sub



