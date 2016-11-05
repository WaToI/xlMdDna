namespace xlMdDna {

	using ExcelDna.Integration;
	using Microsoft.Office.Interop.Excel;
	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using System.IO;
	using System.Reflection;
	using System.Text;
	using System.Text.RegularExpressions;
	using System.Windows.Forms;

	public static class xlMarkdown {
		private static DirectoryInfo MyDoc { get { return new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)); } }
		private static DirectoryInfo saveDir { get { return new DirectoryInfo($@"{MyDoc.FullName}\xlMdDna"); } }
		private static Microsoft.Office.Interop.Excel.Application xl = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
		private static Workbook wb;
		private static Worksheet ws;
		private static Range rng;
		private static ExcelReference caller;
		private static IntPtr userWebBrowser;
		private static IntPtr prevWebBrowser;
		private static string shapName = "";
		private static bool initEnd = false;
		private static string md;
		private static string dq = "\"";

		private static string GetResource(string resourceName) {
			using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
			using (StreamReader reader = new StreamReader(stream))
				return reader.ReadToEnd();
		}

		private static string HTML {
			//get{ return GetResource("xlMdDna.content.def.html"); }
			get {
				return @"
<!DOCTYPE html>
<html lang='ja'>

<head>
	<meta charset='utf-8'>
	<meta http-equiv='X-UA-Compatible' content='IE=edge,chrome=1'>
	<meta name='viewport' content='width=device-width, initial-scale=1'>
	{JQJS}
	{MDJS}
	{MMJS}
	{MMCS}
</head>

<body>
	<div id='md'>
{MDSTR}
	</div>
	<div id='preview'>
	</div>
<script>
var renderer = new marked.Renderer();
renderer.code = function (code, language) {
    if(language == 'mermaid')
        return '<pre class="
+ $"{dq}mermaid{dq}" + @">'+code+'</pre>';
    else
        return '<pre><code>'+code+'</code></pre>';
};

$(document).ready(function(){
	var md = $('#md').text()
	$('#md').empty()
	var html = marked(md ,{ renderer: renderer } );
	$('#preview').html(html);
	mermaid.init();
});

{MMIE}
</script>
</body>

</html>
";
			}
		}

		private static string JQJS {
			//get{ return GetResource("xlMdDna.content.jquery.min.js"); }
			get { return @"<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js'></script>"; }
		}

		private static string MDJS {
			//get{ return GetResource("xlMdDna.content.marked.js"); }
			get { return @"<script src='https://cdnjs.cloudflare.com/ajax/libs/marked/0.3.6/marked.min.js'></script>"; }
		}

		private static string MMJS {
			//get{ return GetResource("xlMdDna.content.mermaid.min.js"); }
			get { return @"<script src='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.js'></script>"; }
		}

		private static string MMCS {
			//get{ return GetResource("xlMdDna.content.mermaid.css"); }
			get { return @"<link rel='stylesheet' type='text/css' href='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.css'>"; }
		}

		private static string MMIE {
			get { return @"mermaid.initialize({flowchart:{htmlLabels:false}});"; }
		}

		private static string MYCSS {
			get {
				return @"
	<style type='text/css'> <!--
		body { width: 100%; background-color: #ffffff;}
		table { border-collapse: collapse;  margin: 10px}
		th { padding: 6px; text-align: left; vertical-align: top; color: #333; background-color: #eee; border: 1px solid #b9b9b9; }
		td { padding: 6px; background-color: #fff; border: 1px solid #b9b9b9; }
		.mermaid{ margin-left: 10px; margin-right: auto; }
	--> </style>
		}
";
			}
		}

		//[ExcelFunction(Name = "Markdown", Description = "About xlMdDna")]
		public static string Markdown(dynamic[,] args) {
			initEnd = init();
			caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			wb = (Workbook)xl.ActiveWorkbook;
			ws = (Worksheet)xl.ActiveSheet;
			rng = (Range)ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];
			shapName = $"{wb.Name}_{ws.Name}_{rng.Address[false, false]}";

			var buf = getArgsString(args);
			md = string.Join("\n", buf).Replace("\u00A0", " ");
			try {
				getPreviewWindow(md.Trim(), $"{shapName}.html");
			}
			catch (Exception ex) {
				Clipboard.SetText($"Err: mermaidFail\n{ex.Message}");
				return "NG";
			}

			return "OK";
		}

		private static bool init() {
			if (!initEnd) {
				if (!saveDir.Exists)
					saveDir.Create();
			}

			return true;
		}

		private static IEnumerable<string> getArgsString(object[,] args) {
			var yLen = args.GetLength(0);
			var xLen = args.GetLength(1);
			var line = "";
			var str = "";
			var rgx = new Regex(@"^(\(|\[|\{)");
			for (var y = 0; y < yLen; y++) {
				line = "";
				for (var x = 0; x < xLen; x++) {
					try {
						if ((str = args[y, x].ToString()) != "ExcelDna.Integration.ExcelEmpty")
							line += (rgx.IsMatch(str) ? "" : " ") + str;
					}
					catch (Exception ex) {
						Clipboard.SetText($"Err: ReadCellFail\n{ex.Message}\n{args[y, x]}");
					}
				}
				yield return line;
			}
		}

		private static string sjisToUtf(string sjisStr) {
			Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
			byte[] bytesData = System.Text.Encoding.UTF8.GetBytes(sjisStr);
			Encoding utf8Enc = Encoding.GetEncoding("UTF-8");
			return utf8Enc.GetString(bytesData);
		}

		private static void windCapture() {
			try {
			}
			catch (Exception ex) {
				MessageBox.Show($"Err: {ex.Message}");
			}
		}

		private static void getPreviewWindow(string md, string fileName = "preview.html") {
			var html = HTML
				.Replace("{JQJS}", JQJS)
				.Replace("{MDJS}", MDJS)
				.Replace("{MMJS}", MMJS)
				.Replace("{MMCS}", MMCS)
				.Replace("{MMIE}", MMIE)
				.Replace("{MDSTR}", md)
				.Trim();

			var path = $@"{saveDir}\{fileName}";
			File.WriteAllText(path, html);
			//web.Navigate(path);
			//web.DocumentText = html;
			try {
				var psi = new ProcessStartInfo("explorer.exe", path);
				var p = Process.Start(psi);
			}
			catch (Exception) {
			}
		}
	}
}