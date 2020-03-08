
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace UpgradeDocmosis
{
	public class Converter
	{

		// cs_ and rs_ both use es_ so this track what kind of end
		private static readonly Stack<string> endTag = new Stack<string>();

		// all the operations in an equation - how we look for an equation
		private static char[] equChars = {'+', '-', '*', '/', '%', '=', '!', '&', '|', '<', '>'};

		/// <summary>
		/// Figure out what to do with this to convert it
		/// </summary>
		/// <param name="field"></param>
		/// <returns></returns>
		public static FieldInfo Parse(string field)
		{

			Trap.trap(field.Contains("["));
			Trap.trap(field.Contains("(") && ! field.StartsWith("{dateFormat") && !field.StartsWith("{numFormat"));
			string query = DataInfo.GetQuery(field);

			// ones we can't handle
			if (query.StartsWith("cc_") || query.StartsWith("barcode:") ||
				query.StartsWith("##") || query.StartsWith("/*") ||
				query == "noTableRowAlternate" || query == "list:continue")
			{
				Trap.trap();
				return new FieldInfo(true);
			}

			// ones we don't handle yet.
			if (query.StartsWith("op:") || query.StartsWith("link:") || query.StartsWith("link_")
						|| query.StartsWith("cr_") || query.StartsWith("rr_") || query.StartsWith("er_"))
			{
				Trap.trap();
				return new FieldInfo(true);
			}

			// if tag
			if (field.StartsWith("cs_"))
			{
				endTag.Push("</wr:if>");
				string item = field.Substring(3);
				if (item.StartsWith("$"))
				{
					Trap.trap();
					return new FieldInfo(null, "[" + field + "]", $"<wr:if select='${{{item.Substring(1)}}}'>");
				}
				if (item.StartsWith("{"))
					return new FieldInfo(null, "[" + field + "]", $"<wr:if select='={item.Substring(1, item.Length-2)}'>");
				query = DataInfo.GetQuery(field.Substring(3));
				return new FieldInfo(field.Substring(3), "[" + field + "]", $"<wr:if select='{query}'>");
			}

			// else tag
			if (field == "else")
			{
				Trap.trap();
				return new FieldInfo(null, "[" + field + "]", "<wr:else/>");
			}
			// else if - we don't support that
			if (field.StartsWith("else_"))
			{
				Trap.trap();
				return new FieldInfo(true);
			}

			// end if or end forEach
			if (field.StartsWith("es_"))
				return new FieldInfo(null, "[" + field + "]", endTag.Pop());

			// forEach - both become a forEach
			if (field.StartsWith("rs_") || field.StartsWith("rr_"))
			{
				endTag.Push("</wr:forEach>");
			
				string item = field.Substring(3);
				if (item.StartsWith("$"))
				{
					Trap.trap();
					Program.logFile.WriteLine($"Check the tag {field} - it has iterative data in a variable");
					return new FieldInfo(null, "[" + field + "]", $"<wr:forEach " +
												$"select='${{{item.Substring(1)}}}' var='{item}'>", item);
				}

				// if no step, we're done
				int index = item.IndexOf(':');
				if (index == -1)
				{
					query = DataInfo.GetQuery(item);
					return new FieldInfo(item, "[" + field + "]", $"<wr:forEach " +
												$"select='{query}' var='{item}'>", item);
				}
				// step
				Trap.trap();
				string step = item.Substring(index + 1);
				if (!step.StartsWith("step"))
				{
					Trap.trap();
					return new FieldInfo(true);
				}
				Trap.trap();
				// if not a number (including whatever down is), we don't handle it
				if (!int.TryParse(item.Substring(index + 1), out int numStep))
				{
					Trap.trap();
					return new FieldInfo(true);
				}
				query = DataInfo.GetQuery(item.Substring(0, index));
				return new FieldInfo(item, "[" + field + "]", $"<wr:forEach " +
												$"select='{query}' step='{numStep}' var='{item}'>", item);
			}

			// out tag, set to template so it processes the html
			if (field.StartsWith("html:"))
			{
				Trap.trap();
				query = DataInfo.GetQuery(field.Substring(5));
				return new FieldInfo(field, "[" + field + "]", $"<wr:out select='{query}' type='TEMPLATE'/>");
			}

			// import sub template
			if (field.StartsWith("ref:"))
			{
				Trap.trap();
				query = DataInfo.GetQuery(field.Substring(4));
				return new FieldInfo(field, "[" + field + "]", $"<wr:import select='{query}' type='TEMPLATE'/>");
			}

			// out sub template
			if (field.StartsWith("refLookup:"))
			{
				Trap.trap();
				query = DataInfo.GetQuery(field.Substring(10));
				return new FieldInfo(field, "[" + field + "]", $"<wr:out select='{query}' type='TEMPLATE'/>");
			}

			// set tag or out tag for a var
			if (field.StartsWith("$"))
			{
				Trap.trap();
				string noDollar = field.Substring(1);
				int index = noDollar.IndexOf('=');
				// it's an out tag
				if (index == -1)
				{
					Trap.trap();
					return new FieldInfo(null, "[" + field + "]", $"<wr:out select='${{{noDollar}}}'/>");
				}

				// it's a set tag
				Trap.trap();
				string select = field.Substring(index + 1).Trim();
				if (select == "null")
				{
					Trap.trap();
					return new FieldInfo(null, "[" + field + "]", $"<wr:set var='{noDollar.Substring(0, index)}'/>");
				}

				if (select == "true" || select == "false" || select.StartsWith("'") || char.IsDigit(select[0]))
				{
					Trap.trap();
					return new FieldInfo(null, "[" + field + "]", $"<wr:set " +
																$"select='{select}' var='{noDollar.Substring(0, index)}'/>");
				}

				Trap.trap();
				select = DataInfo.GetQuery(select);
				return new FieldInfo(field, "[" + field + "]", $"<wr:set " +
															$"select='{select}' var='{noDollar.Substring(0, index)}'/>");
			}

			// it's an equation - out tag
			if (field.StartsWith("{") && field.EndsWith("}"))
			{
				// special case for formatting
				if (field.StartsWith("{dateFormat(") || field.StartsWith("{numFormat("))
				{
					string item = field.Substring(field.IndexOf('(') + 1);
					int index = item.LastIndexOf(')');
					if (index != -1)
					{
						item = item.Substring(0, index).Trim();
						index = item.IndexOf(',');
						string format;
						if (index != -1)
						{
							format = item.Substring(index + 1).Trim();
							item = item.Substring(0, index).Trim();
						}
						else format = null;

						string formatProp = format == null ? "" : $" format='{format}'";
						if (item.StartsWith("'") || char.IsDigit(item[0]) || item.IndexOfAny(equChars) != -1)
							return new FieldInfo(null, "[" + field + "]", $"<wr:out " +
																		$"select='={item}'{formatProp} type='DATE'/>");

						item = DataInfo.GetQuery(item);
						return new FieldInfo(item, "[" + field + "]", $"<wr:out " +
																	$"select='{item}'{formatProp} type='DATE'/>");
					}
				}

				Trap.trap();
				return new FieldInfo(null, "[" + field + "]", $"<wr:out select='={field.Substring(1, field.Length-2)}'/>");
			}

			// plain old out tag
			Trap.trap(field.Contains("{"));
			Trap.trap(field.Contains("_"));
			return new FieldInfo(field, "[" + field + "]", $"<wr:out select='{query}'/>");
		}
	}

	public class FieldInfo
	{
		private static readonly Regex rgx = new Regex("[^a-zA-Z0-9 -.]");

		/// <summary>
		/// If the tag needs to push/pop a node in the data.xml
		/// </summary>
		public enum STACK_MODE
		{
			/// <summary>Do nothing.</summary>
			NOTHING,
			/// <summary>Push this node.</summary>
			PUSH,
			/// <summary>Pop the top node.</summary>
			POP
		}

		/// <summary>
		/// We handle this tag.
		/// </summary>
		/// <param name="nodeName">The node to add to the XML.</param>
		/// <param name="displayName">The text to display for the tag. Normally "[NodeName]"</param>
		/// <param name="fullTagText">The full tag text, example: &lt;wr:out select='select * from...' var='dave'/&gt;</param>
		public FieldInfo(string nodeName, string displayName, string fullTagText)
		{
			if (!string.IsNullOrEmpty(nodeName))
			{
				nodeName = rgx.Replace(nodeName, "_");
				NodeName = nodeName.Split(new char[] {'.'}, StringSplitOptions.RemoveEmptyEntries);
			}

			DisplayName = displayName;
			FullTagText = fullTagText;
		}

		/// <summary>
		/// We handle this tag.
		/// </summary>
		/// <param name="nodeName">The node to add to the XML.</param>
		/// <param name="displayName">The text to display for the tag. Normally "[NodeName]"</param>
		/// <param name="fullTagText">The full tag text, example: &lt;wr:out select='select * from...' var='dave'/&gt;</param>
		/// <param name="varName">For a tag that uses var=, this is the var name.</param>
		public FieldInfo(string nodeName, string displayName, string fullTagText, string varName) 
			: this(nodeName, displayName, fullTagText)
		{
			VarName = varName;
		}

		/// <summary>Do not process this tag</summary>
		public FieldInfo(bool doNotProcess)
		{
			DoNotProcess = doNotProcess;
		}

		/// <summary>
		/// If the tag needs to push/pop a node in the data.xml
		/// </summary>
		public STACK_MODE Mode
		{
			get
			{
				if (FullTagText.Contains("/wr:forEach"))
					return STACK_MODE.POP;
				if (FullTagText.Contains("wr:forEach"))
					return STACK_MODE.PUSH;
				return STACK_MODE.NOTHING;
			}
		}

		/// <summary>
		/// Leave this alone, can't convert
		/// </summary>
		public bool DoNotProcess { get; }
		
		/// <summary>
		/// The node to add to the XML.
		/// </summary>
		public string [] NodeName { get; }

		/// <summary>
		/// The text to display for the tag. Normally "[NodeName]"
		/// </summary>
		public string DisplayName { get; }

		/// <summary>
		/// The full tag text, example: <wr:out select='select * from...' var='dave'/>
		/// </summary>
		public string FullTagText { get; }

		/// <summary>
		/// For a tag that uses var=, this is the var name.
		/// </summary>
		public string VarName { get; }
	}

	/// <summary>
	/// This is for each level we dive in in the data (ie an iterative loop)
	/// </summary>
	public class DataInfo
	{
		/// <summary>
		/// Create the object.
		/// </summary>
		/// <param name="nodeOn">The node to AppendChild() data to when this is the top of the stack</param>
		/// <param name="name">If this is a data/name/field node, use this. Do not use if Var is non-null</param>
		/// <param name="var">If set then select is ${var.field}</param>
		public DataInfo(XmlElement nodeOn, string name, string var)
		{
			NodeOn = nodeOn;
			Name = name;
			Var = var;
		}

		/// <summary>
		/// The node to AppendChild() data to when this is the top of the stack
		/// </summary>
		public XmlElement NodeOn { get; }

		/// <summary>
		/// If this is a data/name/field node, use this. Do not use if Var is non-null
		/// </summary>
		public string Name { get; }

		/// <summary>
		/// If set then select is ${var.field}
		/// </summary>
		public string Var { get; }

		/// <summary>
		/// The query to get data at this level.
		/// </summary>
		/// <param name="field">The field to return.</param>
		/// <returns>The full query for the field.</returns>
		public static string GetQuery(string field)
		{
			DataInfo info = Program.stack.Peek();
			if (!string.IsNullOrEmpty(info.Var))
				return $"${{{info.Var}.{field}}}";

			StringBuilder buf = new StringBuilder();
			foreach (DataInfo item in Program.stack)
				buf.Insert(0, "/" + item.Name);

			string[] nodes = field.Split(new char[] {'.'}, StringSplitOptions.RemoveEmptyEntries);
			foreach (string nodeOn in nodes)
				buf.Append("/" + nodeOn);
			return buf.ToString();
		}
	}
}
