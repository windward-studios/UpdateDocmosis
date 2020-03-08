using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace UpgradeDocmosis
{
	/// <summary>
	/// Upgrade a Docmosis template to Windward
	/// </summary>
	public class Program
	{
		private static XmlDocument xmlDoc;
		public static Stack<DataInfo> stack;
		public static StreamWriter logFile;

		/// <summary>
		/// Convert a Docmosis template to a Windward template, creating a dummy data file and logging what was done.
		/// </summary>
		/// <param name="args">docmosis.docx windward.docx data.xml logging.txt</param>
		static void Main(string[] args)
		{
			if (args.Length != 4)
			{
				Console.Out.WriteLine("Usage: docmosis.docx windward.docx data.xml logging.txt");
				return;
			}

			string srcFilename = Path.GetFullPath(args[0]);
			if (!File.Exists(srcFilename))
			{
				Console.Out.WriteLine($"File {srcFilename} does not exist.");
				return;
			}

			// we do everything on the destination file
			string destFilename = Path.GetFullPath(args[1]);
			if (File.Exists(destFilename))
				File.Delete(destFilename);
			else if (!Directory.Exists(Path.GetDirectoryName(destFilename)))
				Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

			// make sure data filename is ok
			string dataFilename = Path.GetFullPath(args[2]);
			if (File.Exists(dataFilename))
				File.Delete(dataFilename);
			else if (!Directory.Exists(Path.GetDirectoryName(dataFilename)))
				Directory.CreateDirectory(Path.GetDirectoryName(dataFilename));

			xmlDoc = new XmlDocument();
			XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
			xmlDoc.InsertBefore(xmlDeclaration, null);
			XmlElement data = xmlDoc.CreateElement("data");
			xmlDoc.AppendChild(data);
			stack = new Stack<DataInfo>();
			stack.Push(new DataInfo(data, "data", null));

			// and the log file
			string textFilename = Path.GetFullPath(args[3]);
			if (File.Exists(textFilename))
				File.Delete(textFilename);
			else if (!Directory.Exists(Path.GetDirectoryName(textFilename)))
				Directory.CreateDirectory(Path.GetDirectoryName(textFilename));

			using (logFile = new StreamWriter(textFilename))
			{
				logFile.WriteLine($"source: {srcFilename}; destination: {destFilename}; data example: {dataFilename}");

				// create & open the dest document - so we can revise it.
				File.Copy(srcFilename, destFilename);
				using (WordprocessingDocument doc = WordprocessingDocument.Open(destFilename, true))
				{
					ParseDocument(doc.MainDocumentPart.Document);
					foreach (HeaderPart header in doc.MainDocumentPart.HeaderParts)
						ParseDocument(header.Header);
					doc.Save();
				}
			}

			xmlDoc.Save(dataFilename);
			Console.Out.WriteLine($"all done, template={destFilename}; data={dataFilename}; log={textFilename}");
		}

		private static void ParseDocument(OpenXmlElement document)
		{
			// we need to walk paragraphs both because the <<field>> may be across several runs and we are going to change the
			// para contents by removing text and inserting a field
			foreach (Paragraph para in document.Descendants<Paragraph>())
			{
				if (!para.Descendants<Text>().Any(text => text.Text.Contains("<<")))
					continue;
				ParseParagraph(para);
			}

			// if we didn't get any - log it
			foreach (Text text in document.Descendants<Text>())
				if (text.Text.Contains("<<"))
					logFile.WriteLine($"Unmatched << at location {text.Text}");
		}

		private static void ParseParagraph(Paragraph para)
		{
			// after this a run may have multiple fields, but a field is completely in a single Text
			CoalesceRuns(para);

			// now convert
			ConvertToFields(para);
		}

		/// <summary>
		/// Coalesce to a single run per field. If << and >> are not in the same run after this, we ignore them.
		/// </summary>
		/// <param name="para">The para to coalesce</param>
		private static void CoalesceRuns(OpenXmlElement para)
		{
			Run runTextStart = null;
			for (OpenXmlElement element = para.FirstChild; element != null; element = element.NextSibling())
			{
				if (element is ProofError || element is BookmarkStart || element is BookmarkEnd)
					continue;

				Run run = element as Run;
				if (run == null)
				{
					Trap.trap(runTextStart == null && ! (element is ParagraphProperties));
					Trap.trap(runTextStart != null);
					runTextStart = null;
					continue;
				}

				Text text = run.GetFirstChild<Text>();
				if (text == null)
				{
					Trap.trap();
					continue;
				}

				// may be the start (or all) of one or more <<fields>>
				// may also be in runs outside of fields.
				if (runTextStart == null)
				{
					runTextStart = GetTextStart(text);
					continue;
				}

				// if at the end, we need to roll up to one
				int endField = text.Text.IndexOf(">>");
				if (endField == -1)
					continue;

				StringBuilder buf = new StringBuilder();
				OpenXmlElement elemField = runTextStart;
				while (true)
				{
					if (elemField is Run runOn)
						buf.Append(runOn.GetFirstChild<Text>().Text);

					bool isFirst = elemField == runTextStart;
					bool isLast = elemField == element;
					OpenXmlElement elemNext = elemField.NextSibling();
					if (!isFirst)
						elemField.Remove();
					if (isLast)
						break;
					elemField = elemNext;
				}

				// the run we keep now has all the text that we combined
				Text textStart = runTextStart.GetFirstChild<Text>();
				textStart.Text = buf.ToString();

				// restart on this as it may be "<< ... >> ... <<"
				element = runTextStart;
				runTextStart = GetTextStart(textStart);
				Trap.trap(runTextStart != null);
			}
		}

		private static Run GetTextStart(Text text)
		{
			int start = 0;
			while (true)
			{
				start = text.Text.IndexOf("<<", start);
				if (start == -1)
					return null;

				int end = text.Text.IndexOf(">>", start);
				if (end == -1)
					return (Run) text.Parent;

				start = end;
			}
		}


		private static void ConvertToFields(Paragraph para)
		{
			for (OpenXmlElement element = para.FirstChild; element != null; element = element.NextSibling())
			{
				Run run = element as Run;
				if (run == null)
					continue;
				Text text = run.GetFirstChild<Text>();
				if ((!text.Text.Contains("<<")) || (! text.Text.Contains(">>")))
					continue;

				// we will create new runs and insert them, then at the end remove this run
				int offset = 0;
				while (true)
				{
					int start = text.Text.IndexOf("<<", offset);
					if (start == -1)
						break;
					int end = text.Text.IndexOf(">>", start);
					if (end == -1)
						break;
					// do we have text to add?
					if (start > offset)
					{
						Run runText = (Run) run.CloneNode(true);
						runText.GetFirstChild<Text>().Text = text.Text.Substring(offset, start - offset);
						para.InsertBefore(runText, run);
					}

					string field = text.Text.Substring(start + 2, end - start - 2).Trim();
					FieldInfo info = Converter.Parse(field);

					// if we can't handle it, leave it alone
					if (info.DoNotProcess)
					{
						Trap.trap();
						logFile.WriteLine($"Not converting field {field}");
						Run runText = (Run)run.CloneNode(true);
						runText.GetFirstChild<Text>().Text = text.Text.Substring(offset, end + 2 - offset);
						para.InsertBefore(runText, run);
						offset = end + 2;
						continue;
					}

					logFile.WriteLine($"Converting field {field} to tag {info.FullTagText}");

					// add the data
					if (info.NodeName != null && info.NodeName.Length > 0)
					{
						XmlElement root = stack.Peek().NodeOn;
						foreach (string nodeOn in info.NodeName)
						{
							XmlElement item = root.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == nodeOn).FirstOrDefault();
							if (item != null)
								root = item;
							else
							{
								XmlElement data = xmlDoc.CreateElement(nodeOn);
								root.AppendChild(data);
								root = data;
							}
						}

						// need to push/pop on the stack
						if (info.Mode == FieldInfo.STACK_MODE.PUSH)
							stack.Push(new DataInfo(root, info.NodeName[0], info.VarName));
					}
					if (info.Mode == FieldInfo.STACK_MODE.POP) 
						stack.Pop();

					// now add the field
					Run runFieldBegin = new Run();
					runFieldBegin.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Begin });
					para.InsertBefore(runFieldBegin, run);

					Run runFieldInstr = new Run();
					FieldCode fieldCode = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
					fieldCode.Text = $" AUTOTEXTLIST   \\t \"{info.FullTagText}\"  \\* MERGEFORMAT ";
					runFieldInstr.AppendChild(fieldCode);
					para.InsertAfter(runFieldInstr, runFieldBegin);

					Run runFieldSeparate = new Run();
					runFieldSeparate.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Separate });
					para.InsertAfter(runFieldSeparate, runFieldInstr);

					Run runDisplay = new Run();
					Text textDisplay = new Text(info.DisplayName);
					runDisplay.AppendChild(textDisplay);
					para.InsertAfter(runDisplay, runFieldSeparate);

					Run runFieldEnd = new Run();
					runFieldEnd.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.End });
					para.InsertAfter(runFieldEnd, runDisplay);

					offset = end + 2;
				}

				// set to the previous one (and then it will next on it)
				element = run.PreviousSibling() ?? para.FirstChild;

				// if nothing left, remove the run. If something left, reduce it to that.
				if (offset >= text.Text.Length)
				{
					element = run.PreviousSibling();
					run.Remove();
					if (element == null)
						element = para.FirstChild;
				}
				else
					text.Text = text.Text.Substring(offset);
			}
		}
	}
}
