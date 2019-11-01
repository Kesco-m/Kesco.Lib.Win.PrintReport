using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Web.Services.Protocols;

namespace Kesco.Lib.Win.PrintReport
{
	public class Report : IReport
	{
		public event PrintEventHandler EndPrint;

		private RS.ReportingService rs;
		private byte[][] m_renderedReport;
		private Graphics.EnumerateMetafileProc m_delegate;
		private MemoryStream m_currentPageStream;
		private Metafile m_metafile;
		private int m_numberOfPages;
		private bool landscape;
		private int m_currentPrintingPage;
		private int m_lastPrintingPage;

		private int pheight;
		private int pwidth;
		private Margins ma;

		public Report(string url)
		{
			// Create proxy object and authenticate
            Console.WriteLine("{0}: Authenticating to the Web service...", DateTime.Now.ToString("HH:mm:ss fff"));
			try
			{
				rs = new RS.ReportingService(url);
				rs.Credentials = CredentialCache.DefaultCredentials;
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		public byte[][] RenderReport(string reportPath, int id, short paperSize)
		{
			return RenderReport(reportPath, id, "emf", paperSize, false);
		}

		public byte[][] RenderReport(string reportPath, int id, string docType, short paperSize, bool save)
		{
			// Private variables for rendering
            Console.WriteLine("{0}: Start render", DateTime.Now.ToString("HH:mm:ss fff"));
			string deviceInfo = null;
			const string format = "IMAGE";
			Byte[] firstPage = null;
			string encoding;
			string mimeType;
			RS.Warning[] warnings = null;

			RS.ParameterValue[] rp = new RS.ParameterValue[2];
			rp.SetValue(new RS.ParameterValue { Name = "id", Value = id.ToString() }, 0);
			rp.SetValue(new RS.ParameterValue { Name = "DT", Value = DateTime.Now.ToString() }, 1);

			RS.ParameterValue[] reportHistoryParameters = null;
			string[] streamIDs = null;
			Byte[][] pages = null;

			// Build device info based on the start page
			deviceInfo = String.Format(@"<DeviceInfo><OutputFormat>{0}</OutputFormat>", docType);
			if(save)
				deviceInfo += @"<DpiX>300</DpiX><DpiY>300</DpiY>";
			rs.ItemNamespaceHeaderValue = new RS.ItemNamespaceHeader();
			rs.ItemNamespaceHeaderValue.ItemNamespace = RS.ItemNamespaceEnum.PathBased;
            Console.WriteLine("{0}: Get properties", DateTime.Now.ToString("HH:mm:ss fff"));
			RS.Property[] properties = rs.GetProperties(reportPath, null);

			foreach(RS.Property property in properties)
			{
				switch(property.Name.ToLower())
				{
					case "pageheight":
					case "pagewidth":
					case "topmargin":
					case "bottommargin":
					case "rightmargin":
					case "leftmargin":
						deviceInfo +=
							String.Format(@"<{0}>{1}mm</{0}>", property.Name, property.Value);
						break;
				}
			}
			deviceInfo += "</DeviceInfo>";

            Console.WriteLine("{0}: Get first page", DateTime.Now.ToString("HH:mm:ss fff"));
			//Exectute the report and get page count.
			try
			{
				// Renders the first page of the report and returns streamIDs for 
				// subsequent pages
				firstPage = rs.Render(
					reportPath,
					format,
					null,
					deviceInfo,
					rp,
					null,
					null,
					out encoding,
					out mimeType,
					out reportHistoryParameters,
					out warnings,
					out streamIDs);
				// The total number of pages of the report is 1 + the streamIDs         
				m_numberOfPages = streamIDs.Length + 1;
				pages = new Byte[m_numberOfPages][];
                Console.WriteLine("{0}: mast be pages: {1}", DateTime.Now.ToString("HH:mm:ss fff"), m_numberOfPages);
				// The first page was already rendered
				pages[0] = firstPage;

				for(int pageIndex = 1; pageIndex < m_numberOfPages; pageIndex++)
				{
					// Build device info based on start page
					deviceInfo = String.Format(@"<DeviceInfo><OutputFormat>{0}</OutputFormat><StartPage>{1}</StartPage></DeviceInfo>",
							docType, pageIndex + 1);
					pages[pageIndex] = rs.Render(
						reportPath,
						format,
						null,
						deviceInfo,
						rp,
						null,
						null,
						out encoding,
						out mimeType,
						out reportHistoryParameters,
						out warnings,
						out streamIDs);
				}
			}
			catch(SoapException ex)
			{
				Kesco.Lib.Log.Logger.WriteEx(new Kesco.Lib.Log.DetailedException(reportPath, ex, id.ToString(), true));
				Console.WriteLine(ex.Detail.InnerXml);
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				if(pages != null)
					Console.WriteLine("{0}: Number of pages: {1}", DateTime.Now.ToString("HH:mm:ss fff"), pages.Length);
			}
			return pages;
		}

		public bool PrintReport(string printerName, string reportPath, int id, int printID, short paperSize, short copiesCount)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
			this.RenderedReport = this.RenderReport(reportPath, id, paperSize);
			if(null == this.RenderedReport)
				return false;
			try
			{
				// Wait for the report to completely render.
				if(m_numberOfPages < 1)
					return false;
				PrintDocument pd = new PrintDocument();
				rs.ItemNamespaceHeaderValue = new RS.ItemNamespaceHeader();
				rs.ItemNamespaceHeaderValue.ItemNamespace = RS.ItemNamespaceEnum.PathBased;
				RS.Property[] properties = rs.GetProperties(reportPath, null);
				pheight = pd.DefaultPageSettings.PaperSize.Height;
				pwidth = pd.DefaultPageSettings.PaperSize.Width;
				double theight = 0;
				double twidth = 0;
				ma = new Margins(0, 0, 0, 0);
				bool size = false;
				foreach(RS.Property property in properties)
				{
					switch(property.Name.ToLower())
					{
						case "pageheight":
							theight = (double.Parse(property.Value) / 0.254);
							pheight = (int)System.Math.Round(theight);
							size = true;
							break;
						case "pagewidth":
							twidth = double.Parse(property.Value) / 0.254;
							pwidth = (int)System.Math.Round(twidth);
							size = true;
							break;
						case "topmargin":
							ma.Top = (int)(double.Parse(property.Value) / 0.254);
							break;
						case "bottommargin":
							ma.Bottom = (int)(double.Parse(property.Value) / 0.254);
							break;
						case "rightmargin":
							ma.Right = (int)(double.Parse(property.Value) / 0.254);
							break;
						case "leftmargin":
							ma.Left = (int)(double.Parse(property.Value) / 0.254);
							break;
					}
					//Console.WriteLine(property.Name + ": " + property.Value);
				}
                Console.WriteLine("{0}: paper change", DateTime.Now.ToString("HH:mm:ss fff"));

				if(!size)
				{
					if(this.m_currentPageStream != null)
					{
						this.m_currentPageStream.Close();
						this.m_currentPageStream = null;
					}
					m_currentPageStream = new MemoryStream(this.m_renderedReport[0]);
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
					pheight = m_metafile.Height;
					pwidth = m_metafile.Width;
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

				pd.DocumentName =  "?docviewprint=" + id.ToString() + "&docviewtypeid=" + printID.ToString() + "&id=" + id.ToString();
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
				if(null != this.m_renderedReport)
				{
					for(int i = 0; i < this.m_renderedReport.Length; ++i)
					{
						this.m_renderedReport[i] = null;
					}
					this.m_renderedReport = null;
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
				int width = m_metafile.Width;
				int height = m_metafile.Height;
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

		public byte[][] RenderedReport
		{
			get
			{
				return m_renderedReport;
			}
			private set
			{
				m_renderedReport = value;
			}
		}

		private void pd_EndPrint(object sender, PrintEventArgs e)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
			if(this.EndPrint != null)
				EndPrint.BeginInvoke(sender, e, null, null);
		}
	}
}