Option Explicit
  Dim t As table 'テーブルオブジェクト
  Dim region As Range

  Sub insert_general_forms()
    Set t = New table
    Set region = Cells(3, 1).CurrentRegion 'テーブル範囲の定義
    region.Select
    
    setPreferance_forms
    Call t.readRangeH(region)
    renderSQL (0)
    Call t.report

  End Sub
  
  Sub insert_master_forms()
    Set t = New table
    setPreferance_forms
    'テーブル範囲が固定
    Set region = Range(Cells(3.1), Cells(43, 2))
    Call t.readRangeV(region)    'カラム軸が横方向
    renderSQL (0)
    Call t.report

  End Sub
  Sub delete_general_forms()
    Set t = New table
    Set region = Cells(3, 1).CurrentRegion 'テーブル範囲の定義
    
    setPreferance_forms
    Call t.readRangeH(region)    'カラム軸が横方向
    renderSQL (1)

  End Sub
  Sub insertSelectedH()
    Dim name As String
    Dim db As String
    
    Set t = New table
    
    'DB名定義
    db = "dbname"
    Call t.setDatabase(db)
    
    'B1セルにテーブル名を記入するものとする
    name = "tableName"
    Call t.setName(name)
    
    'テーブル範囲が固定
    Set region = Selected
    Call t.readRangeV(region)    'カラム軸が横方向
    renderSQL (0)
    Call t.report
  End Sub
  Private Sub setPreferance_forms()
    Dim name As String
    Dim db As String
    
    'DB名定義
    db = "serious"
    Call t.setDatabase(db)
    
    'B1セルにテーブル名を記入するものとする
    name = ActiveSheet.Cells(1, 2).value
    Call t.setName(name)

  End Sub
  Private Sub renderSQL(Optional mode As Integer = 0)
    Dim tCell As Range 'SQL文の書き込み先 targetCell
  
    'テーブル部分の下に、生成したSQLを出力
    Set tCell = Cells(65535, 1).End(xlUp).Offset(2, 0)
    'mode によりSQL種類を分ける
    If (mode = 1) Then
      tCell.value = t.delete(, 1)
    ElseIf (mode = 2) Then
      tCell.value = t.update
    Else
      tCell.value = t.insert
    End If
    tCell.Select
  End Sub
