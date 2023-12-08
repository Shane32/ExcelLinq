// See https://aka.ms/new-console-template for more information
using Shane32.ExcelLinq.Tests.Models;

Console.WriteLine("Hello, World!");



using var stream1 = new System.IO.FileStream("test.xlsx", System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
var tt = new TestFileContext(stream1);
var pp = tt.GetSheet<Class1>();
var xl = new TestFileContext();
using var stream = new System.IO.FileStream("test.csv", System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
var t = xl.ReadCsv<Class1>(stream);


var g = t;

