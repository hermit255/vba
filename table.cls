Option Explicit
  Dim name As String
  Dim database As String
  'column は string の集合
  Dim columns As New Collection
  'records は collection(record) の集合、record は string の集合
  Dim records As New Collection

  '最低限必要な準備は readRange("レンジオブジェクト") で
  'カラム＋レコードとなるセル範囲(カラム横並び)を指定するだけ。
  '任意で setName("テーブル名"),setDatabase("DB名") で追加設定することもできる。
  '準備ができたら、任意のセルに insert や update 等でSQL文を書き込んでやる用途のクラス。

  Function insert()
    Dim SQL As String 'SQL文のベース部分
    Dim SQL2 As String 'SQL分のレコード部分
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
    Dim newSQL As String 'SQLに追加していく一時用の文字列
    Dim SQL As String '最終的に出力するSQL文
    Dim couple As New Collection 'columns をキー、records を value とする集合
    
    For Each record In records
      Set couple = New Collection
      For Each Column In columns
          couple.Add (vConvert(Column, 2) & " = " & vConvert(record(Column)))
      Next Column
      '最初のカラムをPkeyと仮定して、UPDATE文を発行する
      newSQL = "UPDATE " & getTarget & " SET " & printC(couple, 0) & " WHERE " & columns(pKey) & " = " & record(columns(pKey)) & ";"
      SQL = SQL & newSQL
    Next record
    update = SQL

    If (mode = 1) Then Call report
  End Function

  Function delete(Optional pKey As Integer = 1, Optional mode = 0)
    Dim SQL As String 'SQL文のベース部分
    Dim newSQL As String 'SQLに継ぎ足される新パーツ
    Dim record As Collection
    
    '最初のカラムをPkeyと仮定して、DELETE文を発行する
    For Each record In records
      newSQL = "DELETE FROM " & getTarget & " WHERE " & columns(pKey) & " = " & record(pKey) & ";"
      SQL = SQL & newSQL
    Next record
    delete = SQL
    
    If (mode = 1) Then Call report
  End Function

  Private Function printC(ByVal coll As Collection, Optional mode As Integer = 1)
    'カラム、レコードを書き出す機能 printCollection
    'mode = 1 で空白やNullの対応を変化に対応する
    Dim count As Integer
    Dim data As String
    Dim element As Variant
          
    data = "( "
    For Each element In coll
       '2項目以降はカンマ付与
       If (count > 0) Then data = data & ", "
       If (mode = 1) Then 'mode=1 ならレコード様にシングルクォートで囲む
         data = data & vConvert(element)
       ElseIf (mode = 2) Then 'mode = 2 ならカラム用にバッククオートで囲む
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
    'ひとまず選択範囲の取得メソッドを作ったが、作れば他の方式も可
    'mode = 1 の場合はカラムが縦に並ぶ形式の表に対応する。（デフォルトはカラムが横）
    Dim cRan As Range 'カラム名となる最初の行 columnRange
    Dim rRan As Range '作業レコード行 recordRange
    Dim line As Integer
    
    Dim top As Integer
    Dim bottom As Integer
    Dim left As Integer
    Dim right As Integer
    
    Set columns = New Collection
    Set records = New Collection

    '範囲指定がなければ選択範囲をテーブルとして扱う
    If (IsNull(tRan)) Then Set tRan = Selection
    
    top = tRan.row
    bottom = tRan(tRan.count).row
    left = tRan.Column
    right = tRan(tRan.count).Column

    'readRangeV
    If (mode = 1) Then
      Set cRan = Range(Cells(top, left), Cells(bottom, left))
      Set columns = getCol(cRan)
      'レコード部分の読み取り
      For line = (left + 1) To right
        Set rRan = Range(Cells(top, line), Cells(bottom, line))
        records.Add Item:=getCol(rRan), key:=columns(line)
      Next line
    'mode = 0 の場合（readRangeH）
    Else
      Set cRan = Range(Cells(top, left), Cells(top, right))
      Set columns = getRow(cRan)
      'レコード部分の読み取り
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
    'mode = 1 が getColumn 0 がgetRow に対応
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
    '値をSQL用に変換 valueConvert
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
        'セル内の改行を改行コードに変換
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
    '処理完了時のメッセージ
    Dim msg As String
    msg = "カラム " & columns(1) & " ～ " & columns(columns.count) & " まで、全 " _
          & records.count & " 件について書き出しました。"
    MsgBox msg
  End Sub




