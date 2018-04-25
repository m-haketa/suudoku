Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim iSheet As Worksheet
Dim oSheet As Worksheet

Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = "Z\n14"
'メインロジック
'この関数から処理を開始する
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
'書式設定、画面表示準備など
'ロジックには直接関係なし
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
'※この関数だけmasu,row,column,section,countが値渡しになっていることに注意！
'返り値
'全ての欄が埋まった場合：TRUE
'失敗したのでやり直す場合：FALSE
  
  If setX > 0 And setY > 0 And setN > 0 Then
'指定された欄に指定された値をセット
    Call setNum(masu, row, column, section, count, setY, setX, setN)
  End If

'途中経過を画面表示
  oSheet.Range("A1:I9").Value = masu
  
'ウェイト
  Sleep 10
  
'未記入マスが0ならパズルが解けたことになる。返り値をTRUEにして終了。
  If count = 0 Then
    Check = True
    Exit Function
  Else
    Check = False
  End If
  
  Dim Y As Long
  Dim X As Long
  Call getNextField(masu, row, column, section, Y, X)
  
'数字を１～９まであてはめるロジック
  Dim N As Long
  Dim ret As Boolean
  For N = 1 To 9
    If canSetNum(row, column, section, Y, X, N) Then
      ret = Check(masu, row, column, section, count, Y, X, N)
      If ret Then
'全ての欄が埋まった状態なので返り値をTRUEにして抜ける
        Check = True
        Exit Function
      End If
    End If
  Next

'全ての値を試してもダメだった場合は返り値FALSEで抜ける
End Function


Sub getNextField(ByRef masu, ByRef row, ByRef column, ByRef section, ByRef retY As Long, ByRef retX As Long)
'次に処理を行う欄に対応するY,Xを取得する
'返り値はないが、その代わりにYとXを書き換えて返す。
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
'y行x列にnumの値が入力可能か判定する
'行(row)、列(column)、3x3の枠(section)のそれぞれで、使用済みの数字をビット管理しているため
'該当する行・列・3x3の枠の数値の「ビットor」をとると、使用済みの数字がわかる。
'その「ビットor」の結果と 2^num　の「ビットand」をとることで入力可能かどうかが判定できる。
'
'返り値
'指定したマス(y,x)に指定の数字(num)を入力可能な場合　TRUE
'入力不可能な場合　FALSE

  Dim Check As Long
  Check = (row(Y) Or column(X) Or section(Int((Y + 2) / 3), Int((X + 2) / 3))) And 2 ^ num
  
  If Check = 0 Then
    canSetNum = True
  Else
    canSetNum = False
  End If

End Function


Sub setNum(ByRef masu, ByRef row, ByRef column, ByRef section, ByRef count, ByRef Y As Long, ByRef X As Long, ByRef num)
'返り値はないが、その代わりにmasu,count,row,column,sectionを書き換える。
'
'指定した値（num）を masu(y,x)に書き込み、未記入マス数(count)を１減らす
'さらに、行(row)、列(column)、3x3の枠(section)のそれぞれについて、使用済みの数字を記録する。
'
'row、column、sectionでは、1～9までの数値が使用済かどうかをビットで管理をしていて、
' nが使用済の場合、 2^n　のビットを立てている。
'たとえば、 1行目で1,2,5が使用済の場合には、row(1)= 2^1+2^2+2^5 = 38 となる。
'※2ビット表現だと「0000100110」　※一番右の桁は未使用

  masu(Y, X) = num
  count = count - 1
  
  row(Y) = row(Y) Or 2 ^ num
  column(X) = column(X) Or 2 ^ num
  section(Int((Y + 2) / 3), Int((X + 2) / 3)) = section(Int((Y + 2) / 3), Int((X + 2) / 3)) Or 2 ^ num
  
End Sub

