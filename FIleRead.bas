Public Sub �t�@�C���̑���̎��s()

  '�t�@�C�����J��
  Dim FileAddress As String
  FileAddress = �t�@�C���_�C�A���O���J��()
  
  '�ǂݍ��݃o�b�t�@��p��
  Dim buf As String

  ' �f�[�^�i�[�R���N�V������p��
  Dim inputData As Collection
  Set inputData = New Collection
  
  '�t�@�C�����J��
  Open FileAddress For Input As #1
  
  '�t�@�C������
  Do Until EOF(1)
        '1�s���o�b�t�@�ɓ����
        Line Input #1, buf
        ' �f�[�^�i�[�R���N�V�����ɓ����
        inputData.Add (CStr(buf))
  Loop
  '�t�@�C�������
  Close #1    ''1�Ԃ̃t�@�C������܂�

End Sub
    
    
Private Function �t�@�C���_�C�A���O���J��() As String

    Dim result As Variant

    result = Application.GetOpenFilename( _
            Title:="�e�L�X�g��I�����Ă�������", _
            MultiSelect:=True)

�t�@�C���_�C�A���O���J�� = result(1)

         
End Function