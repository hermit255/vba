Attribute VB_Name = "Table"
Option Explicit
Sub MakeTable()
Attribute MakeTable.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' セルに記入されたテーブル構造を、シェイプとして生成する
'
  'On Error GoTo Warning
  
  Dim Table As Range 'SelectedRange
    Set Table = Selection
  Dim FirstCell As Range
    Set FirstCell = Table.Cells(1)
  Dim LastCell As Range
    Set LastCell = Table.Cells(Table.Count)
  Dim Count As Integer
  
  Dim Left As Single
  Dim Top As Single
  Dim Width As Single
    Width = 150
  Dim Height As Single
    Height = 15
  Dim HeightSum As Single
    HeightSum = 0
  
  Dim SRange As ShapeRange

  
  '選択範囲は一列のみ可
  If Table.Cells(1).Column <> LastCell.Column Then Err.Raise 17
  
  For Count = 1 To Table.Count
    If Count = 1 Then
      Call MakeHeader(Table.Cells(Count), FirstCell.Left, FirstCell.Top, Width, Height)
    ElseIf Count = 2 Then
      Call MakePrimaryKey(Table.Cells(Count), FirstCell.Left, FirstCell.Top, Width, Height, HeightSum)
    Else
      Call MakeColumns(Table.Cells(Count), FirstCell.Left, FirstCell.Top, Width, Height, HeightSum)
    End If
    'グループリストに追加
    If Count > 1 Then SRange.Select False
    Set SRange = Selection.ShapeRange
    '高さ合計を加算
    HeightSum = HeightSum + Height
    '最終行でフレームを生成
    If Count = Table.Count Then
      Call MakeFrame(Table.Cells(1), FirstCell.Left, FirstCell.Top, Width, HeightSum)
      Selection.ShapeRange.ZOrder msoSendToBack
      'グループリストに追加
      SRange.Select False
      Set SRange = Selection.ShapeRange
    End If
  Next Count
  SRange.Group
  ActiveSheet.Activate
  Exit Sub
Warning:
    MsgBox "エラー番号:" & Err.Number & vbCrLf & "テーブル元になるセル(1列のみ可)を選択した状態で実行してください。"
End Sub
Private Sub MakeFrame(Cell, Left, Top, Width, Height)
  
  
  ActiveSheet.Shapes.AddShape(msoShapeRectangle, Left, Top, Width, Height).Select
    With Selection.ShapeRange

      With .Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.75
      End With

      With .Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
     End With

    End With
End Sub
Private Sub MakeHeader(Cell, Left, Top, Width, Height)
    
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, Left, Top, Width, Height).Select
  Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Cell.Value

  With Selection.ShapeRange
   
    With .Line
      .Visible = msoTrue
      .ForeColor.ObjectThemeColor = msoThemeColorText1
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = 0
      .Transparency = 0
    End With
     
    With .Fill
      .ForeColor.ObjectThemeColor = msoThemeColorBackground1
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = -0.150000006
      .Transparency = 0.5
    End With
     
    With .TextFrame2
      .VerticalAnchor = msoAnchorMiddle
    End With
    With .TextFrame2.TextRange.Characters().Font
      .NameComplexScript = "+mn-cs"
      .NameFarEast = "+mn-ea"
      .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
      .Fill.ForeColor.TintAndShade = 0
      .Fill.ForeColor.Brightness = 0
      .Fill.Transparency = 0
      .Fill.Solid
      .Size = 11
      .Name = "+mn-lt"
    End With
    
   End With
End Sub
Private Sub MakePrimaryKey(Cell, Left, Top, Width, Height, Optional AdjustT = 0)
  
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, Left, Top + AdjustT, Width, Height).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Cell.Value

  With Selection.ShapeRange(1)
      With .Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
     End With
    
    With .Fill
      .Visible = msoFalse
    End With

    With .TextFrame2
      .VerticalAnchor = msoAnchorMiddle
    End With
    With .TextFrame2.TextRange.Characters().Font
       .NameComplexScript = "+mn-cs"
       .NameFarEast = "+mn-ea"
       .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
       .Fill.ForeColor.TintAndShade = 0
       .Fill.ForeColor.Brightness = 0
       .Fill.Transparency = 0
       .Fill.Solid
       .Size = 11
       .Name = "+mn-lt"
    End With
    
  End With
End Sub
Private Sub MakeColumns(Cell, Left, Top, Width, Height, Optional AdjustT = 0)
  
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, Left, Top + AdjustT, Width, Height).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Cell.Value

  With Selection.ShapeRange
    With .Line
      .Visible = msoFalse
    End With
    
    With .Fill
      .Visible = msoFalse
    End With
    
    With .TextFrame2
      .VerticalAnchor = msoAnchorMiddle
    End With
    With .TextFrame2.TextRange.Characters().Font
      .NameComplexScript = "+mn-cs"
      .NameFarEast = "+mn-ea"
      .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
      .Fill.ForeColor.TintAndShade = 0
      .Fill.ForeColor.Brightness = 0
      .Fill.Transparency = 0
      .Fill.Solid
      .Size = 11
      .Name = "+mn-lt"
    End With

  End With
End Sub

    

