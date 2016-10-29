namespace xlMdDna {

	using ExcelDna.Integration;
	using Microsoft.Office.Core;
	using Microsoft.Office.Interop.Excel;
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.IO;
	using System.Text;
	using System.Text.RegularExpressions;
	using System.Windows.Forms;

	public static class xlMermaid {
		private static DirectoryInfo MyDoc { get { return new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)); } }
		private static DirectoryInfo saveDir { get { return new DirectoryInfo($@"{MyDoc.FullName}\xlMdDna"); } }
		private static Microsoft.Office.Interop.Excel.Application xl = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
		private static Workbook wb;
		private static Worksheet ws;
		private static Range rng;
		private static ExcelReference caller;
		private static string shapName = "";
		private static bool initEnd = false;
		private static Form wind;
		private static WebBrowser web;
		private static int width = 400 * 2;
		private static int height = 400 * 2;
		private static bool firstTime = true;
		private static dynamic lastStyle = "zoom:300%;";
		private static dynamic lastPos = null;
		private static string md;

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
{MMOP}
</script>
</body>
</html>
";

		private static string mmieop { get { return @"
mermaid.initialize({flowchart:{htmlLabels:false}});
"; } }

		[ExcelFunction(Name = "Mermaid", Description = "About xlMdDna")]
		public static string Mermaid(dynamic[,] args) {
			initEnd = init();
			caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			wb = (Workbook)xl.ActiveWorkbook;
			ws = (Worksheet)xl.ActiveSheet;
			rng = (Range)ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];
			shapName = $"{wb.Name}_{ws.Name}_{rng.Address[false,false]}";

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

				wind = new Form();
				wind.Text = "press {Enter} to Save";
				wind.Width = width;
				wind.Height = height;
				wind.AutoScaleMode = AutoScaleMode.Font;
				//wind.AutoSize = true;
				wind.BackColor = Color.White;
				wind.TopMost = true;
				web = new WebBrowser();
				web.Visible = true;
				web.Dock = DockStyle.Fill;
				wind.Controls.Add(web);
				wind.Show();

				web.DocumentCompleted += (s, e) => {
					if (firstTime) {
						firstTime = false;
						var x = (int)(web.Document.Window.Size.Width / 2 * 2);
						var y = 0;// (int)(web.Document.Window.Size.Height/2*.5);
						web.Document.Body.Style = "zoom:350%;";
						web.Document.Window.ScrollTo(x, y);
					}
				};

				web.PreviewKeyDown += (s, e) => {
					if (e.KeyData == Keys.Enter) {
						windCapture();
						wind.Close();
						web.Dispose();
					}
				};

				wind.FormClosing += (s, e) => {
					firstTime = true;
					initEnd = false;
					//windCapture();
					//e.Cancel = true;
					//wind.Hide();
				};
			}
			if (!wind.Visible) {
				wind.Show();
			}
			wind.FormBorderStyle = FormBorderStyle.Sizable;
			web.ScrollBarsEnabled = true;

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
				wind.FormBorderStyle = FormBorderStyle.None;
				web.ScrollBarsEnabled = false;
				lastPos = web.Document.Window.Position;
				//wind.Activate();
				saveSvg(shapName + ".svg");
				SendKeys.SendWait("%{PRTSC}");
				var bmp = (Bitmap)Clipboard.GetImage();
				bmp.MakeTransparent(Color.White);
				//Clipboard.SetImage(bmp);
				var path = $"{saveDir.FullName}/{shapName}.png";
				bmp.Save(path);
				//xl.ActiveSheet.Paste();
				try {
					var tshap = ws.Shapes.Item(shapName);
					tshap.Delete();
				}
				catch (Exception) { }
				try {
					var shap = ws.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0f, 0f, bmp.Width, bmp.Height);
					shap.Name = shapName;
					shap.Left = float.Parse($"{rng.Offset[0, 1].Left}");
					shap.Top = float.Parse($"{rng.Top}");
				}
				catch (Exception) { }
				wind.FormBorderStyle = FormBorderStyle.Sizable;
				web.ScrollBarsEnabled = true;
			}
			catch (Exception ex) {
				MessageBox.Show($"Err: {ex.Message}");
			}
		}

		private static string getSvg() {
			var pv = web.Document.GetElementById("preview");
			var svgStr = pv.InnerHtml;
			var dq = "\"";
			svgStr = Regex.Replace(svgStr, $@"[^\s]*={dq}{dq}", "");
			svgStr = Regex.Replace(svgStr, $@" *NS\d*:ns\d*:", "");
			svgStr = Regex.Replace(svgStr, $@"NS[^>]*(/*)>", "$1>");
			svgStr = Regex.Replace(svgStr, $@"{dq}(space=)", $"{dq} xmlns:$1");
			svgStr = Regex.Replace(svgStr, $@"<tspan[^>]*><", "<");
			svgStr = Regex.Replace(svgStr, $@" xml(ns)*:[^ ]*", "");
			svgStr = Regex.Replace(svgStr, $@"\s+", " ");
			svgStr = svgStr.Replace("/* */", "");
			//Clipboard.SetText(svgStr);
			return svgStr;
		}

		private static void saveSvg(string fileName = "preview.svg") {
			File.WriteAllText($"{saveDir.FullName}/{fileName}", getSvg());
		}

		private static Form getPreviewWindow(string md, string fileName = "preview.html") {
			var html = defhtml
				.Replace("{MMSTR}", md)
				//.Replace("{MMCSS}", mmcss)
				//.Replace("{MMJS}", mmjs)
				.Replace("{MMOP}", md.StartsWith("graph") ? mmieop : "")
				.Trim();

			var path = $@"{saveDir}\{fileName}";
			File.WriteAllText(path, html);
			//web.Navigate(path);
			web.DocumentText = html;
			return wind;
		}
	}
}