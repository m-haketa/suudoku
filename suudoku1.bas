Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim iSheet As Worksheet
Dim oSheet As Worksheet

Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = "Z\n14"
'���C�����W�b�N
'���̊֐����珈�����J�n����
  Set iSheet = ThisWorkbook.Worksheets(1)
  Set oSheet = ThisWorkbook.Worksheets(2)
  
  Dim masu As Variant
  masu = iSheet.Range("A1:I9").Value
   
  Dim count As Long
  count = 9 * 9
  
  Dim X As Long
  Dim Y As Long
  
  Dim row(1 To 9) As Long
  Dim column(1 To 9) As Long
  Dim section(1 To 3, 1 To 3) As Long
  
  For Y = 1 To 9
    For X = 1 To 9
      If masu(Y, X) <> "" Then
        Call setNum(masu, row, column, section, count, Y, X, masu(Y, X))
      End If
    Next
  Next
 
  Call init(masu)

  Call Check(masu, row, column, section, count, 0, 0, 0)

End Sub


Sub init(ByRef masu)
'�����ݒ�A��ʕ\�������Ȃ�
'���W�b�N�ɂ͒��ڊ֌W�Ȃ�
  iSheet.Range("A1:I9").Font.ColorIndex = 5
  iSheet.Range("A1:I9").Font.Bold = False
  
  Dim X As Long
  Dim Y As Long
  
  For Y = 1 To 9
    For X = 1 To 9
      If masu(Y, X) <> "" Then
        iSheet.Cells(Y, X).Font.ColorIndex = 0
        iSheet.Cells(Y, X).Font.Bold = True
      End If
    Next
  Next

  iSheet.Range("A1:I9").Copy oSheet.Range("A1:I9")
  oSheet.Activate
End Sub


Function Check(ByVal masu, ByVal row, ByVal column, ByVal section, ByVal count, setY As Long, setX As Long, setN) As Boolean
'�����̊֐�����masu,row,column,section,count���l�n���ɂȂ��Ă��邱�Ƃɒ��ӁI
'�Ԃ�l
'�S�Ă̗������܂����ꍇ�FTRUE
'���s�����̂ł�蒼���ꍇ�FFALSE
  
  If setX > 0 And setY > 0 And setN > 0 Then
'�w�肳�ꂽ���Ɏw�肳�ꂽ�l���Z�b�g
    Call setNum(masu, row, column, section, count, setY, setX, setN)
  End If

'�r���o�߂���ʕ\��
  oSheet.Range("A1:I9").Value = masu
  
'�E�F�C�g
  Sleep 10
  
'���L���}�X��0�Ȃ�p�Y�������������ƂɂȂ�B�Ԃ�l��TRUE�ɂ��ďI���B
  If count = 0 Then
    Check = True
    Exit Function
  Else
    Check = False
  End If
  
  Dim Y As Long
  Dim X As Long
  Call getNextField(masu, row, column, section, Y, X)
  
'�������P�`�X�܂ł��Ă͂߂郍�W�b�N
  Dim N As Long
  Dim ret As Boolean
  For N = 1 To 9
    If canSetNum(row, column, section, Y, X, N) Then
      ret = Check(masu, row, column, section, count, Y, X, N)
      If ret Then
'�S�Ă̗������܂�����ԂȂ̂ŕԂ�l��TRUE�ɂ��Ĕ�����
        Check = True
        Exit Function
      End If
    End If
  Next

'�S�Ă̒l�������Ă��_���������ꍇ�͕Ԃ�lFALSE�Ŕ�����
End Function


Sub getNextField(ByRef masu, ByRef row, ByRef column, ByRef section, ByRef retY As Long, ByRef retX As Long)
'���ɏ������s�����ɑΉ�����Y,X���擾����
'�Ԃ�l�͂Ȃ����A���̑����Y��X�����������ĕԂ��B
  retY = 0
  retX = 0

  Dim X As Long
  Dim Y As Long
  For Y = 1 To 9
    For X = 1 To 9
      If IsEmpty(masu(Y, X)) Then
        retY = Y
        retX = X
        Exit Sub
      End If
    Next
  Next

End Sub


Function canSetNum(ByRef row, ByRef column, ByRef section, ByRef Y As Long, ByRef X As Long, ByRef num) As Boolean
'y�sx���num�̒l�����͉\�����肷��
'�s(row)�A��(column)�A3x3�̘g(section)�̂��ꂼ��ŁA�g�p�ς݂̐������r�b�g�Ǘ����Ă��邽��
'�Y������s�E��E3x3�̘g�̐��l�́u�r�b�gor�v���Ƃ�ƁA�g�p�ς݂̐������킩��B
'���́u�r�b�gor�v�̌��ʂ� 2^num�@�́u�r�b�gand�v���Ƃ邱�Ƃœ��͉\���ǂ���������ł���B
'
'�Ԃ�l
'�w�肵���}�X(y,x)�Ɏw��̐���(num)����͉\�ȏꍇ�@TRUE
'���͕s�\�ȏꍇ�@FALSE

  Dim Check As Long
  Check = (row(Y) Or column(X) Or section(Int((Y + 2) / 3), Int((X + 2) / 3))) And 2 ^ num
  
  If Check = 0 Then
    canSetNum = True
  Else
    canSetNum = False
  End If

End Function


Sub setNum(ByRef masu, ByRef row, ByRef column, ByRef section, ByRef count, ByRef Y As Long, ByRef X As Long, ByRef num)
'�Ԃ�l�͂Ȃ����A���̑����masu,count,row,column,section������������B
'
'�w�肵���l�inum�j�� masu(y,x)�ɏ������݁A���L���}�X��(count)���P���炷
'����ɁA�s(row)�A��(column)�A3x3�̘g(section)�̂��ꂼ��ɂ��āA�g�p�ς݂̐������L�^����B
'
'row�Acolumn�Asection�ł́A1�`9�܂ł̐��l���g�p�ς��ǂ������r�b�g�ŊǗ������Ă��āA
' n���g�p�ς̏ꍇ�A 2^n�@�̃r�b�g�𗧂ĂĂ���B
'���Ƃ��΁A 1�s�ڂ�1,2,5���g�p�ς̏ꍇ�ɂ́Arow(1)= 2^1+2^2+2^5 = 38 �ƂȂ�B
'��2�r�b�g�\�����Ɓu0000100110�v�@����ԉE�̌��͖��g�p

  masu(Y, X) = num
  count = count - 1
  
  row(Y) = row(Y) Or 2 ^ num
  column(X) = column(X) Or 2 ^ num
  section(Int((Y + 2) / 3), Int((X + 2) / 3)) = section(Int((Y + 2) / 3), Int((X + 2) / 3)) Or 2 ^ num
  
End Sub

