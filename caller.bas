Attribute VB_Name = "caller"
' ���[�N�u�b�N���J�����̃C�x���g
Private Sub Workbook_Open()
    
    ' txt�ɏ����Ă���O�����C�u������ǂݍ���
    load_from_conf ".\libdef.txt"
       
End Sub



' -------------------- ���W���[���ǂݍ��݂Ɋւ���֐� --------------------



' �ݒ�t�@�C���ɏ����Ă���O�����C�u������ǂݍ��݂܂��B
Sub load_from_conf(conf_path)
    
    ' �S���W���[�����폜
    clear_modules
    
    ' ��΃p�X�ɕϊ�
    conf_path = abs_path(conf_path)
    If Dir(conf_path) = "" Then
        MsgBox "�O�����C�u������`" & conf_path & "�����݂��܂���B"
        Exit Sub
    End If
    
    ' �ǂݎ��
    fp = FreeFile
    Open conf_path For Input As #fp
    Do Until EOF(fp)
        ' �P�s����
        Line Input #fp, temp_str
        If Len(temp_str) > 0 Then
            module_path = abs_path(temp_str)
            If Dir(module_path) = "" Then
                ' �G���[
                MsgBox "���W���[��" & module_path & "�͑��݂��܂���B"
                Exit Do
            Else
                ' ���W���[���Ƃ��Ď�荞��
                include module_path
            End If
        End If
    Loop
    Close #fp

    ThisWorkbook.Save
    
End Sub


' ���郂�W���[�����O������ǂݍ��݂܂��B
' �p�X��.�Ŏn�܂�ꍇ�́C���΃p�X�Ɖ��߂���܂��B
Sub include(file_path)
    ' ��΃p�X�ɕϊ�
    file_path = abs_path(file_path)
    
    ' �W�����W���[���Ƃ��ēo�^
    ThisWorkbook.VBProject.VBComponents.Import file_path
End Sub


' �S���W���[�������������܂��B
Private Sub clear_modules()
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type = 1 Then
            ' ���̕W�����W���[�����폜
            ThisWorkbook.VBProject.VBComponents.Remove component
        End If
    Next component
End Sub


' �t�@�C���p�X���΃p�X�ɕϊ����܂��B
Function abs_path(file_path)

    ' ��΃p�X�ɕϊ�
    If Left(file_path, 1) = "." Then
        file_path = ThisWorkbook.Path & Mid(file_path, 2, Len(file_path) - 1)
    End If
    
    abs_path = file_path

End Function