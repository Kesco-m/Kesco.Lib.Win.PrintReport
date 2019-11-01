using System;
using System.Net;
using Kesco.Lib.Win.PrintReport.RS2005;
using System.Drawing.Printing;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Kesco.Lib.Win.PrintReport
{
	public class Report2005 : IReport
	{
		ReportExecutionService res;

		private MemoryStream m_currentPageStream;
		private Metafile m_metafile;
		private int m_numberOfPages;
		private bool landscape;
		private int m_currentPrintingPage;
		private int m_lastPrintingPage;
		private Graphics.EnumerateMetafileProc m_delegate;

		private int pheight;
		private int pwidth;
		private ExecutionInfo ei;

		public Report2005(string url)
		{
            Console.WriteLine("{0}: Authenticating to the Web service...", DateTime.Now.ToString("HH:mm:ss fff"));
			try
			{
				res = new ReportExecutionService(url);
				res.Credentials = CredentialCache.DefaultCredentials;
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		public event System.Drawing.Printing.PrintEventHandler EndPrint;

		public bool PrintReport(string printerName, string reportPath, int id, int printID, short paperSize, short copiesCount)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
			this.RenderedReport = this.RenderReport(reportPath, id, paperSize);
			if(null == this.RenderedReport)
				return false;
			try
			{

				PrintDocument pd = new PrintDocument();
				//rs.ItemNamespaceHeaderValue = new ItemNamespaceHeader();
				//rs.ItemNamespaceHeaderValue.ItemNamespace = ItemNamespaceEnum.PathBased;
				//Property[] properties = res.GetProperties(reportPath, null);
				pheight = pd.DefaultPageSettings.PaperSize.Height;
				pwidth = pd.DefaultPageSettings.PaperSize.Width;
				double theight = 0;
				double twidth = 0;
				Margins ma = new Margins(0, 0, 0, 0);
				bool size = false;
				if(ei != null)
				{
					ma.Top = (int)(ei.ReportPageSettings.Margins.Top / 0.254);
					ma.Bottom = (int)(ei.ReportPageSettings.Margins.Bottom / 0.254);
					ma.Right = (int)(ei.ReportPageSettings.Margins.Right / 0.254);
					ma.Left = (int)(ei.ReportPageSettings.Margins.Left / 0.254);
					theight = (ei.ReportPageSettings.PaperSize.Height / 0.254);
					pheight = (int)System.Math.Round(theight);
					size = true;

					twidth = ei.ReportPageSettings.PaperSize.Width / 0.254;
					pwidth = (int)System.Math.Round(twidth);
				}

                //Console.WriteLine("{0}: paper change", DateTime.Now.ToString("HH:mm:ss fff"));

				if(!size)
				{
					if(this.m_currentPageStream != null)
					{
						this.m_currentPageStream.Close();
						this.m_currentPageStream = null;
					}
					m_currentPageStream = new MemoryStream(this.RenderedReport[0]);
					// Set its postion to start.
					m_currentPageStream.Position = 0;
					// Initialize the metafile
					if(null != m_metafile)
					{
						m_metafile.Dispose();
						m_metafile = null;
					}
					// Load the metafile image for this page
					m_metafile = new Metafile((Stream)m_currentPageStream);


					pheight = (int)(m_metafile.Height / 300 * 96);
					pwidth = (int)(m_metafile.Width / 300 * 96);
				}

				landscape = false;
				if(pwidth > pheight && !pd.DefaultPageSettings.Landscape)
					landscape = true;

				PrinterSettings printerSettings = new PrinterSettings();
				printerSettings.MaximumPage = m_numberOfPages;
				printerSettings.MinimumPage = 1;
				printerSettings.PrintRange = PrintRange.SomePages;
				printerSettings.FromPage = 1;
				printerSettings.ToPage = m_numberOfPages;
				printerSettings.Copies = copiesCount;
				m_currentPrintingPage = 1;
				m_lastPrintingPage = m_numberOfPages;
				printerSettings.PrinterName = printerName;
				pd.PrinterSettings = printerSettings;

				if(landscape)
				{
					if(pd.DefaultPageSettings.PaperSize.Width != pheight || pd.DefaultPageSettings.PaperSize.Height != pwidth)
					{
						PaperSize papers = new PaperSize(reportPath, pheight, pwidth);
						papers.PaperName = "ReportPrintingLandscape";
						pd.DefaultPageSettings.PaperSize = papers;
					}
					pd.DefaultPageSettings.Landscape = true;
				}
				else
				{
					if(pd.DefaultPageSettings.PaperSize.Width != pwidth || pd.DefaultPageSettings.PaperSize.Height != pheight)
					{
						PaperSize papers = new PaperSize(reportPath, pwidth, pheight);
						papers.PaperName = "ReportPrinting";
						pd.DefaultPageSettings.PaperSize = papers;
					}
					pd.DefaultPageSettings.Landscape = false;
				}

				pd.OriginAtMargins = true;
				pd.DefaultPageSettings.Margins = ma;
				pd.PrintPage += pd_PrintPage;

                pd.DocumentName = "?docviewprint=" + id.ToString() + "&docviewtypeid=" + printID.ToString() + "&id=" + id.ToString();
				pd.EndPrint += pd_EndPrint;

				// Print report
                Console.WriteLine("{0}: Printing report...", DateTime.Now.ToString("HH:mm:ss fff"));
				if(pd.PrinterSettings.IsValid)
					pd.Print();
				else
                    Console.WriteLine("{0}: Encorrect parameters", DateTime.Now.ToString("HH:mm:ss fff"));
				pd.Dispose();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				// Clean up goes here.
			}
			return true;
		}

		public byte[][] RenderedReport { get; private set; }

		public byte[][] RenderReport(string reportPath, int id, short paperSize)
		{
			return RenderReport(reportPath, id, "EMF", paperSize, false);
		}

		public byte[][] RenderReport(string reportPath, int id, string docType, short paperSize, bool save)
		{
			string historyID = null;
			string deviceInfo = "";
			//@"<DeviceInfo><DpiX>300</DpiX><DpiY>300</DpiY></DeviceInfo> ";
			Byte[] results;
			const string format = "IMAGE";
			string encoding = String.Empty;
			string mimeType = String.Empty;
			string extension = String.Empty;
			Warning[] warnings = null;
			string[] streamIDs = null;

			System.Collections.Generic.List<byte[]> list = new System.Collections.Generic.List<byte[]>();
			try
			{
				ParameterValue[] rp = new ParameterValue[2];
				rp.SetValue(new ParameterValue { Name = "id", Value = id.ToString() }, 0);
				rp.SetValue(new ParameterValue { Name = "DT", Value = DateTime.Now.ToString() }, 1);
				ei = res.LoadReport(reportPath, historyID);
				res.SetExecutionParameters(rp, "ru-ru");
				int page = 0;
				while(true)
				{
					deviceInfo = string.Format(@"<DeviceInfo><OutputFormat>{0}</OutputFormat><PrintDpiX>300</PrintDpiX><PrintDpiY>300</PrintDpiY><StartPage>{1}</StartPage></DeviceInfo>", docType, ++page);
					results = res.Render(format, deviceInfo,
							  out extension, out encoding,
							  out mimeType, out warnings, out streamIDs);
					ExecutionInfo2 ei2 = res.GetExecutionInfo2();
					if(results != null && results.Length > 0)
						list.Add(results);
					else
						break;
				}
				m_numberOfPages = list.Count;
				if(m_numberOfPages > 0)
					return list.ToArray();
				else
					return null;
			}
			catch(Exception ex)
			{
				throw new Kesco.Lib.Log.DetailedException(ex.Message, ex, Kesco.Lib.Log.Priority.ExternalError, string.Format("Parameters - reportPath (URL): {0} DocumentID: {1} DocumentType: {2}", reportPath, id, docType));
			}
		}

		private void pd_PrintPage(object sender, PrintPageEventArgs ev)
		{
			ev.HasMorePages = false;
			if(m_currentPrintingPage > m_lastPrintingPage || !MoveToPage(m_currentPrintingPage, ev.Graphics))
				return;

			ReportDrawPage(ev.Graphics);
			// If the next page is less than or equal to the last page, 
			// print another page.
			if(++m_currentPrintingPage <= m_lastPrintingPage)
				ev.HasMorePages = true;
			else
			{
				if(null != this.RenderedReport)
				{
					this.RenderedReport[0] = null;
					this.RenderedReport = null;
					if(this.m_metafile != null)
					{
						this.m_metafile.Dispose();
						this.m_metafile = null;
					}
					if(this.m_currentPageStream != null)
					{
						this.m_currentPageStream.Close();
						this.m_currentPageStream = null;
					}
				}
			}
		}

		// Method to draw the current emf memory stream 
		private void ReportDrawPage(Graphics g)
		{
			if(null == m_currentPageStream || 0 == m_currentPageStream.Length || null == m_metafile)
				return;
			lock(this)
			{
				// Set the metafile delegate.
				m_metafile.SelectActiveFrame(FrameDimension.Page, m_currentPrintingPage - 1);
				int width = m_metafile.Width * 96 / 300;
				int height = m_metafile.Height * 96 / 300;
				m_delegate = MetafileCallback;
				// Draw in the rectangle
				Point[] points = new Point[3];
				Point destPoint = new Point(0, 0);
				Point destPoint1 = new Point(width, 0);
				Point destPoint2 = new Point(0, height);

				points[0] = destPoint;
				points[1] = destPoint1;
				points[2] = destPoint2;
				g.EnumerateMetafile(m_metafile, points, m_delegate);
				points = null;
				// Clean up
				m_delegate = null;
			}
		}

		private bool MoveToPage(Int32 page, Graphics g)
		{
			// Check to make sure that the current page exists in
			// the array list
			if(null == RenderedReport[m_currentPrintingPage - 1])
				return false;
			// Set current page stream equal to the rendered page
			if(m_currentPageStream != null)
			{
				m_currentPageStream.Close();
				m_currentPageStream = null;
			}

			m_currentPageStream = new System.IO.MemoryStream(this.RenderedReport[m_currentPrintingPage - 1]);
			// Set its postion to start.
			m_currentPageStream.Position = 0;
			// Initialize the metafile
			if(null != m_metafile)
			{
				m_metafile.Dispose();
				m_metafile = null;
			}
			// Load the metafile image for this page
			m_metafile = new Metafile(m_currentPageStream);
			return true;
		}

		private bool MetafileCallback(EmfPlusRecordType recordType, int flags, int dataSize, IntPtr data, PlayRecordCallback callbackData)
		{
			byte[] dataArray = null;
			// Dance around unmanaged code.
			if(data != IntPtr.Zero)
			{
				// Copy the unmanaged record to a managed byte buffer 
				// that can be used by PlayRecord.
				dataArray = new byte[dataSize];
				Marshal.Copy(data, dataArray, 0, dataSize);
			}
			// play the record.      
			m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);
			return true;
		}

		private void pd_EndPrint(object sender, PrintEventArgs e)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
			if(this.EndPrint != null)
				EndPrint.BeginInvoke(sender, e, null, null);
		}
	}
}