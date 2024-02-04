/*
 * This software includes the work that is distributed in the Apache License 2.0  
 * This software includes the work that is distributed in the MIT License  
 */

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

//tabula-sharp 0.1.3
//MIT License
//PM> NuGet\Install-Package Tabula -Version 0.1.3
using Tabula;
using Tabula.Detectors;
using Tabula.Extractors;

//PdfPig 0.1.6
//Apache-2.0
//PM> NuGet\Install-Package PdfPig -Version 0.1.6
using UglyToad.PdfPig;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //引数がない場合は終了
            if (args.Length == 0)
            {
                return;
            }

            //引数のファイルをすべて処理
            Console.WriteLine("表データの抽出処理開始");
            foreach (string file in args)
            {
                //絶対パスの取得
                string filepath = Path.GetFullPath(file);
                Console.WriteLine(Environment.NewLine + filepath + "の処理を開始");

                //ファイルの存在確認
                if (!File.Exists(filepath))//ファイルが存在しない場合
                {
                    Console.WriteLine("ファイルが存在しません。");
                    continue;
                }
                else if ((new FileInfo(filepath)).Length == 0)//ファイルサイズ0の場合
                {
                    Console.WriteLine("ファイルサイズが0です。");
                    continue;
                }
                else if (Path.GetExtension(filepath).ToLower() != ".pdf")//拡張子がpdfでない場合
                {
                    Console.WriteLine("拡張子がpdfではありません。");
                    continue;
                }

                //フォルダ、名前
                string folder = Path.GetDirectoryName(filepath);
                string filename = Path.GetFileNameWithoutExtension(filepath);

                //ファイルを開く
                PdfDocument pdfDocument = null;
                Console.WriteLine("ファイルを開いています。");
                try
                {
                    pdfDocument = PdfDocument.Open(filepath, new ParsingOptions() { ClipPaths = true });
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("ファイルが開けません。セキュリティ設定(編集権限)を確認してください。");
                    continue;
                }

                //ファイルが開ける場合の処理
                if (pdfDocument != null)
                {
                    bool LatticeMode;
                    ObjectExtractor oe = new ObjectExtractor(pdfDocument);
                    int num = pdfDocument.GetPages().Count();

                    for (int mode = 0; mode <= 1; mode++)
                    {
                        LatticeMode = mode == 0;

                        if (LatticeMode)
                        {
                            Console.WriteLine("Lattice modeで出力中...");
                        }
                        else
                        {
                            Console.WriteLine("Stream modeで出力中...");
                        }

                        //出力用
                        System.Text.StringBuilder sb = new System.Text.StringBuilder();

                        //すべてのページで処理
                        for (int i = 1; i <= num; i++)
                        {
                            Console.WriteLine(i.ToString() + "/" + num.ToString() + "ページ目を処理中...");
                            PageArea page = oe.Extract(i);

                            IExtractionAlgorithm ea;
                            List<Table> tables = new List<Table>();

                            if (LatticeMode)
                            {
                                ea = new SpreadsheetExtractionAlgorithm();
                                tables = ea.Extract(page);
                            }
                            else
                            {
                                SimpleNurminenDetectionAlgorithm detector = new SimpleNurminenDetectionAlgorithm();
                                var regions = detector.Detect(page);
                                ea = new BasicExtractionAlgorithm();
                                tables = ea.Extract(page.GetArea(regions[0].BoundingBox));
                            }

                            //すべての表を処理
                            foreach (var table in tables)
                            {
                                //すべての行を処理
                                foreach (var row in table.Rows)
                                {
                                    //すべてのセルを処理
                                    foreach (var cell in row)
                                    {
                                        string cellValue = cell.GetText();
                                        bool doubleQuotationMark = cellValue.Contains(",") || cellValue.Contains("\n") || cellValue.Contains("\r");
                                        if (doubleQuotationMark)
                                        {
                                            sb.Append("\"");
                                        }
                                        sb.Append(cellValue);
                                        if (doubleQuotationMark)
                                        {
                                            sb.Append("\"");
                                        }
                                        sb.Append(",");
                                    }
                                    sb.Append(Environment.NewLine);
                                }
                                sb.Append(Environment.NewLine);
                            }
                        }

                        //ファイルに出力
                        StreamWriter sw = null;
                        string outputfile;
                        if (LatticeMode)
                        {
                            outputfile = folder + @"\" + filename + "_LatticeMode.csv";
                        }
                        else
                        {
                            outputfile = folder + @"\" + filename + "_StreamMode.csv";
                        }

                        try
                        {
                            //ファイルを作成
                            sw = new StreamWriter(outputfile, false, new System.Text.UTF8Encoding(true));
                            sw.Write(sb.ToString());
                            Console.WriteLine(outputfile + "へ出力しました。");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.WriteLine("ファイル出力ができませんでした。");
                        }
                        finally
                        {
                            if (sw != null)
                            {
                                sw.Close();
                            }
                        }
                    }

                    //pdfを閉じる
                    pdfDocument.Dispose();
                }
            }
        }
    }
}