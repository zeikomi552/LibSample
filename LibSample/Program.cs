// See https://aka.ms/new-console-template for more information
using ArgsManagerLib;
using DocumentFormat.OpenXml;
using System.Runtime.CompilerServices;
using TableToJsonlConverter.Conveters;

Console.WriteLine("Hello, World!");


string[] keys = new string[]
    {
        "-type",
        "-file",
        "-scol",
        "-srow",
        "-chcol",
    };


ArgsManager _argsManager = new ArgsManager(keys, args);

string CheckInput(string key, string message)
{
    var val = _argsManager[key];

    if (string.IsNullOrEmpty(val))
    {
        Console.WriteLine(message);
        return Console.ReadLine()!;
    }
    else
    {
        return val;
    }
}

void Init()
{



    var type = CheckInput("-type", "タイプを入力してください。 -> xlsx または csv");
    var file = CheckInput("-file", "読み込むファイルのパスを入力してください。").Replace("\"","");

    switch (type)
    {
        case "xlsx":
            {
                var scol = int.Parse(CheckInput("-scol", "開始列を指定してください(1からはじまる)"));
                var srow = int.Parse(CheckInput("-srow", "開始行を指定してください(1から始まる)"));
                var chcol = int.Parse(CheckInput("-chcol", "チェックする列を指定してください(この列が空を見つけ次第データ取得を終了します)"));
                var sheetno = int.Parse(CheckInput("-sheetno", "取り込むExcelのシート番号を指定してください(0から始まる)"));

                ZkExcelToJsonl zkExcelToJsonl = new ZkExcelToJsonl(file, true, scol, srow, chcol, sheetno);
                zkExcelToJsonl.Read();
                var header = zkExcelToJsonl.GetHeader();
                zkExcelToJsonl.Write("test.jsonl");
                break;
            }
        case "csv":
            {
                ZkCsvToJsonl zkCsvToJsonl = new ZkCsvToJsonl(file, System.Text.Encoding.UTF8, true);
                zkCsvToJsonl.Read();
                var header = zkCsvToJsonl.GetHeader();
                zkCsvToJsonl.Write("test2.jsonl");
                break;
            }
    }
}

Init();