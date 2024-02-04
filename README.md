# pdftables2csv
 
## 概要

PDFファイル内の表を抽出してcsvファイルに出力します。  

## 使用方法

* PDFファイルを`pdftables2csv.exe`にドラッグ＆ドロップする。または  
* コマンドの実行 `pdftables2csv.exe input.pdf`  

## 出力

2パターンの抽出アルゴリズムで抽出し、上記の場合であれば、`input_LatticeMode.csv`と`input_StreamMode.csv`の2つのファイルを出力します。  

## 動作確認環境
Microsoft Windows10 x64 + .NET Framework 4.8

## ライセンス

This software is released under the MIT License.   
詳細については、[LICENSE](./LICENSE) ファイルを参照してください。  

This software includes the following works:  
[ThirdPartyLICENSEフォルダ](./ThirdPartyLICENSE/)内の全てのファイルを確認してください。  

|Work|Documents|Remarks|
|:----|:----|:----|
|PdfPig 0.1.6|[LICENSE](./ThirdPartyLICENSE/PdfPig/LICENSE)<br>[NOTICE.txt](./ThirdPartyLICENSE/PdfPig/NOTICES.txt)|[NuGet](https://www.nuget.org/packages/PdfPig/0.1.6)|
|tabula-sharp 0.1.3|[LICENSE](./ThirdPartyLICENSE/tabula-sharp/LICENSE)|[NuGet](https://www.nuget.org/packages/Tabula/0.1.3)|
|System.ValueTuple 4.5.0|[LICENSE.txt](./ThirdPartyLICENSE/System.ValueTuple/LICENSE.TXT)|[NuGet](https://www.nuget.org/packages/System.ValueTuple/4.5.0)|

