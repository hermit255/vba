Attribute VB_Name = "use_unit"
Option Explicit
Sub renderAllUnit()
  Dim entire As New Unit
  Dim setting As New Collection
  Dim temp As Collection

  dim top as integer
  dim bottom as integer
  dim nGen as integer

  '処理範囲を設定しているが、select を利用して動的にしてもよい。
  'top = selection(0).row
  'bottom = selection(selection.count).row
  top = 2
  bottom = 10
  ngen = 3

  '世代ごとの設定 temp を1世代の設定として、子世代をまとめる括弧（前後）、
  '子世代要素の区切り文字、自世代を囲む括弧、子要素の列指定（右隣以外の場合）の順に設定
  
  'デフォルトの装飾設定
  Set temp = New Collection
  temp.Add "("
  temp.Add ")"
  temp.Add ","
  temp.Add ""
  temp.Add ""
  temp.Add 0
  setting.Add temp
  '第1世代の装飾設定
  Set temp = New Collection
  temp.Add ""
  temp.Add ""
  temp.Add ","
  temp.Add "【"
  temp.Add "】"
  temp.Add 0
  setting.Add temp
  '第2世代の装飾設定
  Set temp = New Collection
  temp.Add "("
  temp.Add ")"
  temp.Add "/"
  temp.Add ""
  temp.Add ""
  temp.Add 0
  setting.Add temp
  '第0世代は区切り文字だけあれば十分
  Set temp = New Collection
  temp.Add ""
  temp.Add ""
  temp.Add ","
  temp.Add ""
  temp.Add ""
  temp.Add 0
  
  '表全体を表す unit オブジェクトとして entire という
  '第0世代を作成して、処理の起点にする。
  entire.setTop(top)
  entire.setBottom(bottom)
  entire.setPreferance(setting)
  entire.applySetting(temp)
  entire.MakeChildren
  entire.fractal(nGen)
  entire.renderMapping
  MsgBox (entire.mapping)
    
End Sub
