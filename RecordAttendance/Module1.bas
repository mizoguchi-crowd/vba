Attribute VB_Name = "Module1"
Private Const ���t�� As String = "A"
Private Const �o�Ύ����� As String = "B"
Private Const �ދΎ����� As String = "C"

Private Sub Shukkin_Click()
  Call �ō�����(�o�Ύ�����)
End Sub

Private Sub Taikin_Click()
  Call �ō�����(�ދΎ�����)
End Sub

Private Sub �ō�����(�ō���)
  Call ���t�L�^
  Dim �ō��s As String
  �ō��s = �ō��s�擾()
  Dim �ō��敪 As String
  �ō��敪 = �ō��敪�擾(�ō���)
  
  If (���d�ō��`�F�b�N(�ō��s, �ō���)) Then
    MsgBox ("�{���͂��ł�" + �ō��敪 + "���Ԃ�ō��ς̂��߁A�L�^���܂���ł���")
    Exit Sub
  Else
    Sheet1.Cells(�ō��s, �ō���) = ���ݎ����擾()
    MsgBox (�ō��敪 + "���Ԃ�ō����܂���")
  End If
End Sub

Private Function ���ݎ����擾() As String
  ���ݎ����擾 = Format(Now, "hh:mm:ss")
End Function

Private Sub ���t�L�^()
  Dim �������t As String
  �������t = Format(Now, "YYYY/MM/DD")
  Dim ���t��ŏI�s�̓��t As String
  ���t��ŏI�s�̓��t = Sheet1.Cells(Rows.Count, ���t��).End(xlUp)
  
  If (���t��ŏI�s�̓��t = �������t) Then
    '�������Ȃ�
  Else
    Sheet1.Cells(Sheet1.Cells(Rows.Count, ���t��).End(xlUp).Row + 1, ���t��) = �������t
  End If

End Sub

Private Function �ō��s�擾() As String
  �ō��s�擾 = Sheet1.Cells(Rows.Count, ���t��).End(xlUp).Row
End Function

Private Function ���d�ō��`�F�b�N(�ō��s, �ō���) As Boolean
  If (Sheet1.Cells(�ō��s, �ō���) = "") Then
    ���d�ō��`�F�b�N = False
  Else
    ���d�ō��`�F�b�N = True
  End If
End Function

Private Function �ō��敪�擾(�ō���) As String
  If (�ō��� = �o�Ύ�����) Then
    �ō��敪�擾 = "�o��"
  ElseIf (�ō��� = �ދΑō���) Then
    �ō��敪�擾 = "�ދ�"
  End If
  
End Function
