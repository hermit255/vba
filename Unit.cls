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
'�\�̓���J�����ɓ���l�����񂾏ꍇ�ɁA�����̍s��
'�����O���[�v�Ƃ��Ĉ������߂̃N���X�B
Dim key As String
Dim col As Integer
Dim top As Integer
Dim bottom As Integer

Dim setting As Collection '�S�̂̐ݒ�f�[�^
Dim generation As Integer
Dim childCol As Integer
Dim children As New Collection 'unit �I�u�W�F�N�g�̏W��
Public mapping As String 'renderMapping �̏o�͌���

'�����p����
Dim delimiter As String
Dim encP As String 'encloserPre
Dim encS As String 'encloserSur
Dim prefix As String
Dim surfix As String

Public Sub MakeChildren()
  'top ���� bottom �܂ł𑖍����āA�z���ɂȂ� unit �I�u�W�F�N�g children �𐶐�����
  '������A��ꐢ��� unit�i�\�S�́j�I�u�W�F�N�g�͕W�����W���[���ō�邱�ƂɂȂ�B
  Dim child As Object 'children �̗v�f
  Dim ran As Range
  '�q�O���[�v�̃L�[�J�����͈����Őݒ�B�f�t�H���g�͐e�O���[�v�L�[�̉E��
  If childCol = 0 Then
    Set ran = Cells(top, col + 1)
  Else
    Set ran = Cells(top, childCol)
  End If

  '�q�ƂȂ��ɐV�����l���o����A����l�������s�𒲍����ăI�u�W�F�N�g�𐶐�
  '�����e�̖����܂ŌJ��Ԃ�
  Do Until ran.Row > bottom
    Set child = New Unit
    Call child.setKey(ran.Column, ran.Value)
    Call child.setGeneration(generation + 1)
    Call child.setTop(ran.Row)
    '�قȂ�l�܂��͐e�I�u�W�F�N�g�̖��s�̎�O�܂ōs������
    Do Until (ran.Value <> ran.Offset(1, 0).Value) Or (ran.Row = bottom)
      Set ran = ran.Offset(1, 0)
    Loop
    Call child.setBottom(ran.Row)
    'value ���󔒂� unit �Ƃ���ȍ~�̐���͖�������
    if( ran.value = "" ) then children.Add child

    Set ran = ran.Offset(1, 0)
  Loop
End Sub

sub test(lastGen As Integer)
  '������̓��ꃆ�j�b�g���Z����������悤�ȋ@�\�������Ă��������낢����
  Dim child As Unit
end sub

Sub fractal(lastGen As Integer)
  '�w�肵������܂Ń}�b�s���O�f�[�^�̐������J��Ԃ��B
  Dim child As Unit

  For Each child In children
    Call child.setPreferance(setting)
    child.applySetting
    '�w�肵������̎�O�Ŏ����㐶���̃��[�v���X�g�b�v
    If (generation < lastGen - 1) Then
      child.MakeChildren
      child.fractal (lastGen)
    End If
    child.renderMapping
  Next child
End Sub

Public Sub renderMapping()
  '����̓���q�\�����e�L�X�g�o�͂���B���ォ����s����z��B
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
  
  '�����Ŏw�肵���ݒ�f�[�^�܂��͎�����p�̐ݒ�f�[�^������Γǂݍ���
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