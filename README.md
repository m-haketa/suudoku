# suudoku

ロジックを大幅に書き換えた新バージョンを追加しました

- suudoku11.bas
Excel VBA 標準モジュールのソースコードをエクスポートしたファイルです。

- suudoku11.xlsm
上記ソースを組み込んでいるエクセルファイル本体です。

1シート目に問題を入力して、Ctrl+Shift+Zを押す（あるいは、main関数を起動する）と
マクロが動きます。


以下、旧バージョンです

- suudoku1.bas
- suudoku2.bas

Excel VBA 標準モジュールのソースコードをエクスポートしたファイルです。
1枚目のシートのA1からI9セルに問題を入れると、
2枚目のシートのA1からI9セルに解答を表示します。


suudoku1.bas、suudoku2.basの差異は「getNextField」関数の中だけで、他はまったく同じです。

ロジックの考え方そのものは、
たぶん「エクセルの真髄」さんの下記URLと、ほぼ同じだろうと思います。

- suudoku1.bas
 https://excel-ubara.com/excelvba5/EXCELVBA231.html

- suudoku2.bas
 https://excel-ubara.com/excelvba5/EXCELVBA231_2.html

- suudoku.xlsx
　エクセルシートのサンプルです。このファイルに上記ソースコードをインポートすると動作します

