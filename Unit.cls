VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'表の特定カラムに同一値が並んだ場合に、それらの行を
'同じグループとして扱うためのクラス。
Dim key As String
Dim col As Integer
Dim top As Integer
Dim bottom As Integer

Dim setting As Collection '全体の設定データ
Dim generation As Integer
Dim childCol As Integer
Dim children As New Collection 'unit オブジェクトの集合
Public mapping As String 'renderMapping の出力結果

'装飾用文字
Dim delimiter As String
Dim encP As String 'encloserPre
Dim encS As String 'encloserSur
Dim prefix As String
Dim surfix As String

Public Sub MakeChildren()
  'top から bottom までを走査して、配下になる unit オブジェクト children を生成する
  '性質上、第一世代の unit（表全体）オブジェクトは標準モジュールで作ることになる。
  Dim child As Object 'children の要素
  Dim ran As Range
  '子グループのキーカラムは引数で設定可。デフォルトは親グループキーの右隣
  If childCol = 0 Then
    Set ran = Cells(top, col + 1)
  Else
    Set ran = Cells(top, childCol)
  End If

  '子となる列に新しい値が出たら、同一値が続く行を調査してオブジェクトを生成
  'これを親の末尾まで繰り返す
  Do Until ran.Row > bottom
    Set child = New Unit
    Call child.setKey(ran.Column, ran.Value)
    Call child.setGeneration(generation + 1)
    Call child.setTop(ran.Row)
    '異なる値または親オブジェクトの末行の手前まで行を下る
    Do Until (ran.Value <> ran.Offset(1, 0).Value) Or (ran.Row = bottom)
      Set ran = ran.Offset(1, 0)
    Loop
    Call child.setBottom(ran.Row)
    'value が空白の unit とそれ以降の世代は無視する
    if( ran.value = "" ) then children.Add child

    Set ran = ran.Offset(1, 0)
  Loop
End Sub

sub test(lastGen As Integer)
  '同世代の同一ユニットをセル結合するような機能があってもおもしろいかも
  Dim child As Unit
end sub

Sub fractal(lastGen As Integer)
  '指定した世代までマッピングデータの生成を繰り返す。
  Dim child As Unit

  For Each child In children
    Call child.setPreferance(setting)
    child.applySetting
    '指定した世代の手前で次世代生成のループをストップ
    If (generation < lastGen - 1) Then
      child.MakeChildren
      child.fractal (lastGen)
    End If
    child.renderMapping
  Next child
End Sub

Public Sub renderMapping()
  '世代の入れ子構造をテキスト出力する。末代から実行する想定。
  Dim data As String
  Dim child As Variant
  For Each child In children
    If Not (data = "") Then data = data & delimiter
    data = data & child.mapping
  Next child
  If (children.Count = 0) Then
    data = prefix & key & surfix
  Else
    data = prefix & key & surfix & encP & data & encS
  End If
  mapping = data
End Sub

Sub setKey(c As Integer, k As String)
  col = c
  key = k
End Sub

Sub setTop(t As Integer)
  top = t
End Sub

Sub setBottom(b As Integer)
  bottom = b
End Sub

Sub setGeneration(g As Integer)
  generation = g
End Sub

Sub setPreferance(s As Collection)
  Set setting = s
End Sub

Sub applySetting(Optional col As Collection = Null)
  Dim myset As Collection
  
  '引数で指定した設定データまたは自世代用の設定データがあれば読み込み
  If Not col Is Nothing Then
    Set myset = col
  ElseIf (setting.Count >= generation + 1) Then
    Set myset = setting(generation + 1)
  Else
    Set myset = setting(1)
  End If
  
    encP = myset(1)
    encS = myset(2)
    delimiter = myset(3)
    prefix = myset(4)
    surfix = myset(5)
    childCol = myset(6)

End Sub
