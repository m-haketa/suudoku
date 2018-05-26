Option Explicit
Option Base 1

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const ScreenUpdate As Boolean = False

Dim iSheet As Worksheet
Dim oSheet As Worksheet

'盤面を表す配列
Dim Matrix()  As Long

'入力可能数字を表現する配列
'2進数で各数字の状態を表していて、入力不能な場合１、入力可能な場合０が入る
'たとえば、2進数で000011111（＝10進数で31）の場合、9～6は入力可、5～1は入力不可
Dim MatrixFlag() As Long

'MatrixFlagに格納される数値について、2進数表示で1が何個あるかをあらかじめ数えておく
'例　225を2進数表示したときに１のビットが何個あるかは、bitCountArr(225）で算出できる
Dim bitCountArr() As Long



'''以下、履歴関連変数
'次に入力する数値が何個目かを表す番号
'この番号が82になったら、すべてのマスにデータが入ったことになる
Dim History_Count As Long


'入力したx座標、y座標と番号
Dim History_y(1 To 81) As Long
Dim History_x(1 To 81) As Long
Dim History_n(1 To 81) As Long

'上記、x座標、y座標、番号の内容
'0:初期値、1:確定、2:未確定
Dim History_Kind(1 To 81) As Long

Enum kind_type
  Initial = 1
  Fixed = 2
  Pending = 3
End Enum

'上記、x,y,n入力「前」のMatrixFlagの状態
'未確定の場合のみ入力する
'※それ以外の場合は、NULLにしておくか？
Dim History_MatrixFlag(1 To 81) As Variant




'''以下、試行錯誤をする場合の「次に試すマス」の情報
Dim Next_y As Long
Dim Next_x As Long

'（next_y,next_x）のMatrixFlagのコピー
Dim Next_MatrixFlag As Long

'（next_y,next_x）に入力できない数値が何個あるか
Dim Next_UsedNumberCount As Long


Sub main()
'メインロジック
'この関数から処理を開始する
  History_Count = 1
  
  Set iSheet = ThisWorkbook.Worksheets(1)
  Set oSheet = ThisWorkbook.Worksheets(2)
  
  Dim StartTime As Double
  Dim EndTime As Double
  Dim ElapsedTimeString As String
  
  ReDim Matrix(1 To 9, 1 To 9) As Long
  ReDim MatrixFlag(1 To 9, 1 To 9) As Long
  
  Call init

'ここから実質的なスタート
  StartTime = Timer
  Debug.Print StartTime
  
  bitCountArr = createBitArray(9)
    
  Dim MatrixTemp As Variant
  MatrixTemp = iSheet.Range("A1:I9").Value
   
  Dim x As Long
  Dim y As Long
   
  For y = 1 To 9
    For x = 1 To 9
      If MatrixTemp(y, x) <> "" Then
        Call setNum(y, x, CLng(MatrixTemp(y, x)), kind_type.Initial)
      End If
    Next
  Next

  Call parse
  
  EndTime = Timer
  Debug.Print EndTime

  ElapsedTimeString = Format((EndTime - StartTime), "0.0000")
  Debug.Print ElapsedTimeString & "秒経過"
  
  Call display
End Sub


Function parse() As Boolean
  
'''確定できるマス目を入力していく
'入力不可能なマスを発見した場合にはエラーが発生する
'その場合には、即座に処理を中断し、次に進む
  On Error GoTo Break
  Call Check1
  On Error GoTo 0
  
  If History_Count > 81 Then
'すべてのマス目に入力完了
    parse = True
    Exit Function
  End If

'''試行錯誤開始
'入力対象となる数値
  Dim n As Long
  For n = 1 To 9

'numが入力できるかどうかを表すフラグ
    Dim flag As Long
    flag = MatrixFlag(Next_y, Next_x) And 2 ^ (n - 1)
    
    If flag = 0 Then
'暫定的に入力可能な数値を入力
      Call setNum(Next_y, Next_x, n, kind_type.Pending)

'IF文の中で再起呼び出し
      If parse() Then
'すべのマス目に入力完了
        parse = True
        Exit Function
      Else
'入力失敗　元に戻す
        Call back_history
        If ScreenUpdate Then
          Call display
        End If
      End If
    End If
  Next

Break:
  parse = False

End Function


Sub init()
'書式設定、画面表示準備など
'ロジックには直接関係なし
  iSheet.Range("A1:I9").Font.ColorIndex = 5
  iSheet.Range("A1:I9").Font.Bold = False
  
  Dim x As Long
  Dim y As Long
  
  For y = 1 To 9
    For x = 1 To 9
      If iSheet.Cells(y, x) <> "" Then
        iSheet.Cells(y, x).Font.ColorIndex = 0
        iSheet.Cells(y, x).Font.Bold = True
      End If
    Next
  Next

  iSheet.Range("A1:I9").Copy oSheet.Range("A1:I9")
    
  Call display
  
  oSheet.Activate
End Sub


Function Check1() As Boolean
'１．入れられる数値が１つしかないマスがないかをチェック
'２．行、列、3x3に着目して１つのマスにしか入れられない数字がないかをチェックする
'３．試行錯誤するときの候補となすマスを探す

'返り値　１つでもマスを埋めた場合　TRUE、１つも埋められなかった場合　FALSE

  Check1 = False
 
  Dim y As Long
  Dim x As Long
  Dim n As Long

'このマスまでチェックをしたら終了
  Dim lastY As Long
  Dim lastX As Long
  lastY = 9
  lastX = 9
  
'試行錯誤で試すマスも同時に探索する
'まず初期化しておく
  Call init_next
    
   
  Do
   For y = 1 To 9
    For x = 1 To 9
      If lastY = y And lastX = x Then GoTo Break
      
'ブランクチェック
      If Matrix(y, x) > 0 Then GoTo Continue
           
'''１．数値が１つだけしか入れられないマスのチェック
      n = 0
      If bitCountArr(MatrixFlag(y, x)) = 8 Then
'数値が１つだけしか入れられないマスの処理
        n = 511 Xor MatrixFlag(y, x)
        n = WorksheetFunction.Log(n * 2, 2)
      End If
                
'''２．行、列、3x3に着目して１つのマスにしか入れられない数字がないかをチェックする
'縦方向チェック
      If n = 0 Then
        n = Search(1, x, 9, x, y, x)
      End If
      
      If n = 0 Then
'横方向チェック
        n = Search(y, 1, y, 9, y, x)
      End If
        
      If n = 0 Then
'3x3チェック
'数字n は同じ3x3の欄に入れられない
        Dim gy As Long
        Dim gx As Long
        gy = Int((y - 1) / 3) * 3
        gx = Int((x - 1) / 3) * 3
        n = Search(gy + 1, gx + 1, gy + 3, gx + 3, y, x)
      End If
                
      If n = 0 Then
'''上記１，２で、入力すべきマスがなかった場合
'''３．試行錯誤するときの候補となすマスを探す
        Call set_next(y, x, MatrixFlag(y, x))
      
      Else
'入力すべきマスがあった場合（n>0の場合）
        Call setNum(y, x, n, kind_type.Fixed)
        Check1 = True
      
'このマスからさらにもう１周探査する
        lastY = y
        lastX = x
        Call init_next
      End If
        
Continue:
    Next
   Next
  Loop
Break:

End Function



Function Search(y1 As Long, x1 As Long, y2 As Long, x2 As Long, yBase As Long, xBase As Long) As Long
'(y1,x1)～(y2,x2)のうち、(yBase,xBase)以外すべてで入力不可の数字を探す
'返り値：１つの数値しか入力できない場合の、その数値　※そういうものがない場合は0を返す
  Dim y As Long
  Dim x As Long
  Dim temp As Long
  temp = 2 ^ 9 - 1

   For y = y1 To y2
    For x = x1 To x2
'1つだけしか入れられないマスを識別
      If y = yBase And x = xBase Then
'値が入れられるかどうかを判定するマス自身については、マスに値が入れられる（＝bitが0）でないとダメなので、
'ビットを反転させたうえでandを取る
        temp = temp And (511 Xor MatrixFlag(y, x))
      Else
        temp = temp And MatrixFlag(y, x)
      End If
    Next
   Next
   
   If temp > 0 Then
     '通常は、条件を満たす数字は１つしかありえないはず
     Search = WorksheetFunction.Log(temp * 2, 2)
   Else
     Search = 0
   End If
End Function


Sub setMatrixFlag(y1 As Long, x1 As Long, y2 As Long, x2 As Long, n1 As Long, n2 As Long, ParamArray Exclude())
'指定のマス（y1,x1）～(y2,x2)まで、指定の値 n1～n2 が入力できない旨 MatrixFlagを設定する
'ただし、Excludeが指定されている場合には、そのマスを除く　例：（exY1,exX1,exY2,exX2)など、2個1組で指定する

  Dim y As Long
  Dim x As Long
  Dim n As Long
  Dim temp As Long
  
  Dim exCount As Long
  
  For y = y1 To y2
   For x = x1 To x2
    For n = n1 To n2
     For exCount = LBound(Exclude) To UBound(Exclude) Step 2
      If y = Exclude(exCount) And x = Exclude(exCount + 1) Then
        GoTo Continue
      End If
     Next
       temp = MatrixFlag(y, x) Or 2 ^ (n - 1)
       MatrixFlag(y, x) = temp
    Next
Continue:
   Next
  Next
End Sub


Sub setNum(y As Long, x As Long, n As Long, kind As Long)
'指定のマス(y,x)に指定の数字nをセットする（Matrixの値を変える）
'それに連動して、MatrixFlagも適切に変化させる
  
'matrix、matrixflagにデータを追加する前に履歴を追加
  Call push_history(y, x, n, kind)

  Dim c As Long
  Dim temp As Long
  
'元のmatrixに格納
  Matrix(y, x) = n
  
'全ての数字は(y,x)に入れられない
  Call setMatrixFlag(y, x, y, x, 1, 9)
  
'数字n は同じ列に入れられない
  Call setMatrixFlag(1, x, 9, x, n, n)

'数字n は同じ行に入れられない
  Call setMatrixFlag(y, 1, y, 9, n, n)

'数字n は同じ3x3の欄に入れられない
  Dim gy As Long
  Dim gx As Long
  
  gy = Int((y - 1) / 3) * 3
  gx = Int((x - 1) / 3) * 3
  
  Call setMatrixFlag(gy + 1, gx + 1, gy + 3, gx + 3, n, n)

  If ScreenUpdate Then
    Call display
  End If
End Sub


Function createBitArray(bitCount As Long) As Long()
'bitCount　　※配列を準備する最大ビット数（0 to 2^bitcount-1で配列が確保される）
'bitCountArr 1が何ビットあるか判定用の配列。たとえばbitCountArr(511) = 9になる。
  Dim retArr() As Long
  ReDim retArr(0 To 2 ^ 9 - 1) As Long
     
  retArr(0) = 0
  
  Dim I As Long
  Dim c As Long
  Dim Cdiff As Long
  
  For I = 1 To bitCount
   Cdiff = 2 ^ (I - 1)
   For c = 0 To Cdiff - 1
    retArr(c + Cdiff) = retArr(c) + 1
   Next
  Next
  
  createBitArray = retArr
End Function


Sub push_history(y As Long, x As Long, n As Long, kind As Long)
  History_y(History_Count) = y
  History_x(History_Count) = x
  History_n(History_Count) = n
  History_Kind(History_Count) = kind
  
  If (kind = kind_type.Pending) Then
  '入力内容が、未確定の場合には、ロールバックすることを勧化、
  '現時点のMatrixFlagのバックアップを取っておく
    Dim tempArray(9, 9) As Long
    Dim tx As Long
    Dim ty As Long
    
    For ty = 1 To 9
      For tx = 1 To 9
        tempArray(ty, tx) = MatrixFlag(ty, tx)
      Next
    Next
    History_MatrixFlag(History_Count) = tempArray
  End If

  History_Count = History_Count + 1
End Sub


Sub back_history()
'pendingレコードに行き当たるまで、matrixを戻す
  Do
    History_Count = History_Count - 1
    Matrix(History_y(History_Count), History_x(History_Count)) = 0
  Loop Until History_Kind(History_Count) = kind_type.Pending
   
'matrix_flagを復元
  Dim tx As Long
  Dim ty As Long
    
  For ty = 1 To 9
    For tx = 1 To 9
      MatrixFlag(ty, tx) = History_MatrixFlag(History_Count)(ty, tx)
    Next
  Next
        
'next_y、next_xも値を復元
  Next_y = History_y(History_Count)
  Next_x = History_x(History_Count)

'たぶん、支障はないが念のため、matrixflagをEmptyに設定しておく
  History_MatrixFlag(History_Count) = Empty
End Sub


Sub init_next()
'次の候補マスの初期化
  Next_y = 0
  Next_x = 0
  Next_UsedNumberCount = 0
End Sub


Sub set_next(y As Long, x As Long, MatrixFlag As Long)
'次の候補マスを設定する
'もし、入力不能なマスが出てきた場合には、例外を発生させる

  Dim UsedNumberCount As Long
  UsedNumberCount = bitCountArr(MatrixFlag)
  
'入力不能マスがある場合には即座に実行中止
'エラーを発生させる（＝元ルーチンでは次の試行に移る）
  If UsedNumberCount = 9 Then
    Call Err.Raise(99999, "set_next", "raise exception")
  End If
  
'入力不能な数値が一番多いマスの情報を保存しておく
  If Next_UsedNumberCount < UsedNumberCount Then
    Next_y = y
    Next_x = x
    Next_MatrixFlag = MatrixFlag
    Next_UsedNumberCount = UsedNumberCount
  End If

End Sub


Sub display()
'現在の状況を画面に表示する
'ロジックには直接関係なし

'通常の表
  oSheet.Cells(1, 1).Resize(9, 9).Value = Matrix
  
'入力不能な数をビットで表した表
  oSheet.Cells(11, 1).Resize(9, 9).Value = MatrixFlag
  
'入力不能なマスをそれぞれの数字ごとに表示した表×９（K列～AM列まで）
'各マスごとの入力可能な文字を文字列で表した表（A21～I29）　を作成
  Dim numX As Long
  Dim numY As Long
  
  Dim x As Long
  Dim y As Long
  
  Dim num As Long
  
  'MatrixFlagを各数値ごとに分解して格納する表
  Dim SepArray(1 To 9, 1 To 9) As Long
  
  'MatrixFlagのうち入力可能な数値をStringで格納する表
  Dim NotUsedArray(1 To 9, 1 To 9) As String
  
  Dim Sep As Long
  Dim NotUsed As Long
  
  For numY = 0 To 2
   For numX = 0 To 2
    
'表示対象となる数字 num
    num = numY * 3 + numX + 1
    
    For y = 1 To 9
      For x = 1 To 9
        Sep = MatrixFlag(y, x) And 2 ^ (num - 1)
        SepArray(y, x) = Sgn(Sep) * num
      
        NotUsed = (1 - Sgn(Sep)) * num
        If NotUsed > 0 Then
          NotUsedArray(y, x) = NotUsedArray(y, x) & NotUsed
        End If
      Next
    Next
    
    oSheet.Cells(1 + numY * 10, 11 + numX * 10).Resize(9, 9).Value = SepArray
       
   Next
  Next
   
  oSheet.Cells(21, 1).Resize(9, 9).Value = NotUsedArray
  
  Sleep 50
  
End Sub
