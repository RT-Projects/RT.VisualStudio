using System;
using RT.Util.ExtensionMethods;
using RT.Util;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ConvertCommentStyle
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Please specify a file path as a command-line parameter.");
                return 1;
            }

            var text = File.ReadAllLines(args[0]);
            var isComment = text.Select(line => Regex.IsMatch(line, @"^\s*///", RegexOptions.Multiline)).ToArray();
            var result = new List<string>();
            foreach (var gr in isComment.GroupConsecutive())
            {
                if (!gr.Key)
                {
                    result.AddRange(text.Skip(gr.Index).Take(gr.Count));
                    continue;
                }

                var indentationLength = Regex.Match(text[gr.Index], @"^\s*", RegexOptions.Multiline).Length;
                var indentation = new string(' ', indentationLength) + "/// ";
                var wrapWidth = 130 - indentationLength;
                var wrap = Ut.Lambda((string str) => str.Trim().WordWrap(wrapWidth).ToArray());
                var chunk = XElement.Parse("<item>{0}</item>".Fmt(text.Subarray(gr.Index, gr.Count).Select(line => Regex.Replace(line, @"\s*/// ", "", RegexOptions.Multiline)).JoinString(Environment.NewLine)), LoadOptions.PreserveWhitespace);
                var elements = chunk.Elements().ToArray();
                foreach (var elem in elements)
                    process(elem);

                IEnumerable<string> ret;

                // Heuristic: If there is *only* a summary tag, and it fits on one line, make it compact
                string[] wrapped;
                if (elements.Length == 1 && elements[0].Name.LocalName == "summary" && (wrapped = wrap(elements[0].GetContent())).Length == 1)
                    result.Add("{0}<summary>{1}</summary>".Fmt(indentation, wrapped[0]));
                else
                {
                    var retStr = elements.Select(elem => "{0}{3}{1}</{2}>".Fmt(elem.GetTag(), wrap(elem.GetContent()).JoinString(Environment.NewLine).Indent(4), elem.Name.LocalName, Environment.NewLine)).JoinString(Environment.NewLine);
                    result.Add(Regex.Replace(retStr, @"^", indentation, RegexOptions.Multiline));
                }
            }

            var resultStr = result.JoinString(Environment.NewLine);
            var original = text.JoinString(Environment.NewLine);
            if (resultStr != original)
                File.WriteAllText(args[0], resultStr);
            return 0;
        }

        private static void process(XElement elem)
        {
            var leaveNewLines = elem.Elements("para").Any() || elem.Elements("code").Any() || elem.Elements("list").Any();
            foreach (var node in elem.Nodes())
            {
                if (node is XText && !leaveNewLines)
                {
                    var xText = (XText) node;
                    xText.Value = Regex.Replace(xText.Value, @"\r?\n", " ");
                }
                else if (node is XElement)
                {
                    var xElement = (XElement) node;
                    if (xElement.Name.LocalName == "para")
                        process(xElement);
                    else if (xElement.Name.LocalName == "list")
                    {
                        foreach (var subElem in xElement.Elements("item"))
                            foreach (var subElem2 in subElem.Elements("description"))
                            {
                                process(subElem2);
                                subElem2.FirstNode.AddBeforeSelf(new XText(Environment.NewLine + "    "));
                            }
                    }
                }
            }
        }
    }

    static class Extensionification
    {
        public static string GetContent(this XElement element)
        {
            return element.Nodes().Select(n => n.ToString()).JoinString();
        }

        public static string GetTag(this XElement element)
        {
            var sb = new StringBuilder();
            sb.Append("<");
            sb.Append(element.Name.LocalName);
            foreach (var attr in element.Attributes())
                sb.Append(@" {0}=""{1}""".Fmt(attr.Name.LocalName, attr.Value.HtmlEscape()));
            sb.Append(">");
            return sb.ToString();
        }
    }
}
