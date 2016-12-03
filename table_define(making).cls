VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "table_define"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim name As String
Dim columns As New Collection 'columns has property
Dim p As New Collection 'property
Dim SQLelement As New Collection
Sub setName(Optional str As String = "tablename")
  name = str
End Sub
Sub generateSQL()
    Dim SQL As String 'SQL文
    Dim count As Integer
    count = 0

    SQL = "CREATE TABLE `" & name & "` "
    SQL = printC(SQLelement)
    SQL = SQL & " ENGINE = InnoDB;"
    insert = SQL
End Sub
 Function printC(ByVal coll As Collection)
    'カラム、レコードを書き出す機能 printCollection
    Dim count As Integer
    Dim data As String
          
    data = "( "
    For Each element In coll
      '2項目以降はカンマ付与
      If (count > 0) Then data = data & ", "
      data = data & element
      count = count + 1
    Next element
      data = data & " )"
      printC = data
  End Function
  Function getCol(ran As Range)
    Dim data As New Collection
    Dim row As Integer
    Dim val As String
    
    Dim top As Integer
    Dim bottom As Integer
    
    top = ran.Column
    bottom = ran(ran.count).Column
    col = ran.col
    
    For row = top To bottom
      val = Cells(row, col).value
      data.Add val
    Next row
    Set getCol = data
  End Function

