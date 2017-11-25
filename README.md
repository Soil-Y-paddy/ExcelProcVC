# ExcelProcVC
VisualC++でExcelを扱うテスト(改良歓迎)
<dl><
<dt>OleWrp.h / OleWrp.cpp </dt>
<dd> COMオブジェクトを取り扱うためのクラスを定義
<ul>
<li>OleWrap :: OLE操作用クラス</li>
<li>SafeArrayCtrl :: Variant２次元配列操作クラス</li>
<li>VariantCtrl :: Variant型　＜＝＞ プリミティブ型 の相互変換(date型は未実装)
</ul>
</dd>
<dt>ExcelProc.h / ExcelProc.cpp </dt>
<dd> エクセル操作クラス<br>
<ul>
<li>ブック操作 ( 新規作成 / 開く / 閉じる / 名前を付けて保存 )</li>
<li>シート操作 ( 追加(左のみ) / 選択(id, 名前)  )</li>
<li>セル操作   ( Range.Value を取得 / 設定 ) </li>
<li> おまけ : 列番号(1,2,3...) を列名(A,B,C,...Z,AA,AB,..)に変換する
</ul>
※とりあえず必要最小限のみなので、どんどん追加してください。。。<br>
 ★未実装項目
<ol>
<li>複数のブックを開いたときの取扱</li>
<li>シートのコピー、削除、シート名の変更</li>
<li>ブックやシート横断の操作</li>
<li>画像、グラフ、VBAマクロの編集や操作</li>
</ol>
</dd>
