Attribute VB_Name = "use_unit"
Option Explicit
Sub renderAllUnit()
  Dim entire As New Unit
  Dim setting As New Collection
  Dim temp As Collection

  dim top as integer
  dim bottom as integer
  dim nGen as integer

  '�����͈͂�ݒ肵�Ă��邪�Aselect �𗘗p���ē��I�ɂ��Ă��悢�B
  'top = selection(0).row
  'bottom = selection(selection.count).row
  top = 2
  bottom = 10
  ngen = 3

  '���ゲ�Ƃ̐ݒ� temp ��1����̐ݒ�Ƃ��āA�q������܂Ƃ߂銇�ʁi�O��j�A
  '�q����v�f�̋�؂蕶���A��������͂ފ��ʁA�q�v�f�̗�w��i�E�׈ȊO�̏ꍇ�j�̏��ɐݒ�
  
  '�f�t�H���g�̑����ݒ�
  Set temp = New Collection
  temp.Add "("
  temp.Add ")"
  temp.Add ","
  temp.Add ""
  temp.Add ""
  temp.Add 0
  setting.Add temp
  '��1����̑����ݒ�
  Set temp = New Collection
  temp.Add ""
  temp.Add ""
  temp.Add ","
  temp.Add "�y"
  temp.Add "�z"
  temp.Add 0
  setting.Add temp
  '��2����̑����ݒ�
  Set temp = New Collection
  temp.Add "("
  temp.Add ")"
  temp.Add "/"
  temp.Add ""
  temp.Add ""
  temp.Add 0
  setting.Add temp
  '��0����͋�؂蕶����������Ώ\��
  Set temp = New Collection
  temp.Add ""
  temp.Add ""
  temp.Add ","
  temp.Add ""
  temp.Add ""
  temp.Add 0
  
  '�\�S�̂�\�� unit �I�u�W�F�N�g�Ƃ��� entire �Ƃ���
  '��0������쐬���āA�����̋N�_�ɂ���B
  entire.setTop(top)
  entire.setBottom(bottom)
  entire.setPreferance(setting)
  entire.applySetting(temp)
  entire.MakeChildren
  entire.fractal(nGen)
  entire.renderMapping
  MsgBox (entire.mapping)
    
End Sub
