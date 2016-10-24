namespace xlMdDna {

	using ExcelDna.Integration;
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.IO;
	using System.Text;
	using System.Windows.Forms;

	public static class MarkDown {
		private static dynamic xl = ExcelDnaUtil.Application;
		private static bool initEnd = false;
		private static Form wind;
		private static WebBrowser web;
		private static DirectoryInfo MyDoc { get { return new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)); } }

		private static string defhtml = @"
<!DOCTYPE html>
<html lang='ja'>
<head>
<meta charset='utf-8'>
<meta http-equiv='X-UA-Compatible' content='IE=edge,chrome=1'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<script src='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.js'></script>
<link rel='stylesheet' type='text/css' href='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.css'>
</head>
<body style='background-color: #ffffff;'>
<div id='preview' class='mermaid'>
{MMSTR}
</div>
<script>
mermaid.initialize({flowchart:{htmlLabels:false}})
</script>
</body>
</html>
";

		[ExcelFunction(Name = "Mermaid", Description = "About xlMdDna")]
		public static string Mermaid(dynamic[,] args) {
			var buf = getArgsString(args);
			var md = string.Join("\n", buf);
			getPreviewWindow(md.Replace("\u00A0", " "));
			try {
				//web.ScrollBarsEnabled = true;
				wind.FormBorderStyle = FormBorderStyle.Sizable;
				wind.Show();
				wind.Activate();
			}
			finally { }

			return "OK";
		}

		private static IEnumerable<string> getArgsString(object[,] args) {
			var yLen = args.GetLength(0);
			var xLen = args.GetLength(1);
			for (var y = 0; y < yLen; y++) {
				var line = "";
				for (var x = 0; x < xLen; x++) {
					if (args[y, x].ToString() != "ExcelDna.Integration.ExcelEmpty")
						line += args[y, x].ToString();
				}
				yield return line;
			}
		}

		private static string sjisToUtf(string sjisstr) {
			Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
			//string sjisstr = sjisEnc.GetString(loaddata);
			byte[] bytesData = System.Text.Encoding.UTF8.GetBytes(sjisstr);
			Encoding utf8Enc = Encoding.GetEncoding("UTF-8");
			return utf8Enc.GetString(bytesData);
		}

		private static bool init() {
			if (!initEnd) {
				wind = new Form();
				wind.BackColor = Color.White;
				web = new WebBrowser();
				web.Visible = true;
				web.Dock = DockStyle.Fill;
				wind.Controls.Add(web);
				wind.Closing += (s, e) => {
					windCapture();
					e.Cancel = true;
					wind.Hide();
				};
			}
			return true;
		}

		private static void windCapture(string fileName = "preview.png") {
			try {
				saveSvg();
				wind.FormBorderStyle = FormBorderStyle.None;
				//web.ScrollBarsEnabled = false;
				wind.Activate();
				SendKeys.SendWait("%{PRTSC}");
				var bmp = (Bitmap)Clipboard.GetImage();
				bmp.MakeTransparent(Color.White);
				//Clipboard.SetImage(bmp);
				var path = $"{MyDoc.FullName}/{fileName}";
				bmp.Save(path);
				//xl.ActiveSheet.Paste();
				xl.ActiveSheet.Shapes.AddPicture(path, 0, -1, 0, 0, bmp.Width, bmp.Height);
			}
			catch (Exception ex) {
				MessageBox.Show($"Err: {ex.Message}");
			}
		}

		private static void saveSvg(string fileName = "preview.svg") {
			try {
				var pv = web.Document.GetElementById("preview");
				var svgStr = pv.InnerHtml;
				Clipboard.SetText(svgStr);
				File.WriteAllText($"{MyDoc.FullName}/{fileName}", svgStr);
			}
			catch (Exception) { }
		}

		private static Form getPreviewWindow(string md) {
			initEnd = init();
			web.DocumentText =
				defhtml
					.Replace("{MMSTR}", md)
			//.Replace("{MMCSS}", mmcss)
			//.Replace("{MMJS}", mmjs)
			;

			return wind;
		}
	}
}