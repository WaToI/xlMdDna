namespace xlMdDna {

	using ExcelDna.Integration;
	using System;
	using System.Windows.Forms;

	public class App {
		private static Version AppVer = new Version(0, 1, 0);

		[ExcelCommand(MenuName = "xlMdDna", MenuText = "About")]
		public static void About() {
			MessageBox.Show( $@"xlMdDna.	Ver: {AppVer}
thanks
  Excel-DNA: https://excel-dna.net/
  mermaid: https://knsv.github.io/mermaid
");
		}
	}
}