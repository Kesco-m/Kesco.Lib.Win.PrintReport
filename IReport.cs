using System;
namespace Kesco.Lib.Win.PrintReport
{
	public interface IReport
	{
		event System.Drawing.Printing.PrintEventHandler EndPrint;
		bool PrintReport(string printerName, string reportPath, int id, int printID, short paperSize, short copiesCount);
		byte[][] RenderedReport { get; }
		byte[][] RenderReport(string reportPath, int id, short paperSize);
		byte[][] RenderReport(string reportPath, int id, string docType, short paperSize, bool save);
	}
}
