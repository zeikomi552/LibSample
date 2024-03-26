// See https://aka.ms/new-console-template for more information
using ArgsManagerLib;
using DocumentFormat.OpenXml;
using System.Runtime.CompilerServices;
using System.Text;
using TableToJsonlConverter.Conveters;

Console.WriteLine("Hello, World!");


List<KeyValuePair<string, string>> keys = new List<KeyValuePair<string, string>>()
{
    new KeyValuePair<string, string>("-type", "File type -> csv or xlsx or sqlserver"),
    new KeyValuePair<string, string>("-ifile", "input file path"),
    new KeyValuePair<string, string>("-ofile", "output faile path"),
    new KeyValuePair<string, string>("-scol", "start column number(Numbers starting with 1). xlsx type only"),
    new KeyValuePair<string, string>("-srow", "start row number(Numbers starting with 1). xlsx type only"),
    new KeyValuePair<string, string>("-chcol", "check column number(Numbers starting with 1). When it finds an empty cell, it finishes the process"),
    new KeyValuePair<string, string>("-sheet", "sheet no. xlsx type onl"),
    new KeyValuePair<string, string>("-cs", "Connection string. sqlserver type only."),
    new KeyValuePair<string, string>("-sql", "Sql Command. sqlserver type only."),
    new KeyValuePair<string, string>("-enc", "Charactor Encoding. csv type only.(ex.utf-8, Shift_JIS, EUC-JP, iso-2022-jp)"),
};


var keylist = (from x in keys
               select x.Key).ToArray();

ArgsManager _argsManager = new ArgsManager(keylist, args);

string CheckInput(string key, string message)
{
    var val = _argsManager[key];

    if (string.IsNullOrEmpty(val))
    {
        Console.WriteLine(message);
        return Console.ReadLine()!.Replace("\"", "");
    }
    else
    {
        return val;
    }
}

string GetValue(string key)
{
    var tmp = (from x in keys where x.Key.Equals(key) select x).FirstOrDefault().Value;

    if (tmp != null)
    {
        return tmp;
    }
    else
    {
        return string.Empty;
    }
}

void Init()
{
    string key = "-type";
    var type = CheckInput(key, GetValue(key));

    switch (type)
    {
        case "xlsx":
            {
                key = "-scol";
                var scol = int.Parse(CheckInput(key, GetValue(key)));
                key = "-srow";
                var srow = int.Parse(CheckInput(key, GetValue(key)));
                key = "-chcol";
                var chcol = int.Parse(CheckInput(key, GetValue(key)));
                key = "-sheet";
                var sheetno = int.Parse(CheckInput(key, GetValue(key)));
                key = "-ifile";
                var ifile = CheckInput(key, GetValue(key));
                key = "-ofile";
                var ofile = CheckInput(key, GetValue(key));


                ZkExcelToJsonl zkExcelToJsonl = new ZkExcelToJsonl(ifile, true, scol, srow, chcol, sheetno);
                zkExcelToJsonl.Read();
                var header = zkExcelToJsonl.GetHeader();
                zkExcelToJsonl.Write(ofile);
                break;
            }
        case "csv":
            {
                key = "-ifile";
                var ifile = CheckInput(key, GetValue(key));
                key = "-ofile";
                var ofile = CheckInput(key, GetValue(key));
                key = "-enc";
                var enc = CheckInput(key, GetValue(key));

                EncodingProvider provider = System.Text.CodePagesEncodingProvider.Instance;
                var encoding = provider.GetEncoding(enc);

                ZkCsvToJsonl zkCsvToJsonl = new ZkCsvToJsonl(ifile, encoding == null ? Encoding.UTF8 : encoding, true);
                zkCsvToJsonl.Read();
                var header = zkCsvToJsonl.GetHeader();
                zkCsvToJsonl.Write(ofile);
                break;
            }
        case "sqlserver":
            {
                key = "-cs";
                var cs = CheckInput(key, GetValue(key));
                key = "-sql";
                var sql = CheckInput(key, GetValue(key));
                key = "-ofile";
                var ofile = CheckInput(key, GetValue(key));

                ZkSQLServerToJsonl zkCsvToJsonl = new ZkSQLServerToJsonl(cs, sql);
                zkCsvToJsonl.Read();
                var header = zkCsvToJsonl.GetHeader();
                zkCsvToJsonl.Write(ofile);
                break;
            }
    }
}

Init();