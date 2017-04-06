// Multiple files ANSI to utf-8 converter
//http://stackoverflow.com/questions/5172172/multiple-files-ansi-to-utf-8-converter


var indir = "in";
var outdir = "out";
function ansiToUtf8(fin, fout) {
    var ansi = WScript.CreateObject("ADODB.Stream");
    ansi.Open();
    ansi.Charset = "x-ansi";
    ansi.LoadFromFile(fin);
    var utf8 = WScript.CreateObject("ADODB.Stream");
    utf8.Open();
    utf8.Charset = "UTF-8";
    utf8.WriteText(ansi.ReadText());
    utf8.SaveToFile(fout, 2 /*adSaveCreateOverWrite*/);
    ansi.Close();
    utf8.Close();
}
var fso = WScript.CreateObject("Scripting.FileSystemObject");
var folder = fso.GetFolder(indir);
var fc = new Enumerator(folder.files);
for (; !fc.atEnd(); fc.moveNext()) {
    var file = fc.item();
    ansiToUtf8(indir+"\\"+file.name, outdir+"\\"+file.name);
}