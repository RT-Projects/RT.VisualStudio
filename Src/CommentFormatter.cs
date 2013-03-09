using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using RT.Util;
using RT.Util.ExtensionMethods;

namespace RT.VisualStudio
{
    public static class CommentFormatter
    {
        private static string[] _inlineTags = new[] { "see", "paramref", "typeparamref", "c" };
        private static string[] _blockLevelTags = new[] { "code", "para", "list", "description" };

        public static string ReformatComments(string source)
        {
            var text = Regex.Split(source.TrimEnd(), @"\r?\n", RegexOptions.Singleline);
            var isComment = text.Select(line => Regex.IsMatch(line, @"^\s*///", RegexOptions.Multiline)).ToArray();
            var result = new StringBuilder();

            foreach (var gr in isComment.GroupConsecutive())
            {
                if (!gr.Key)
                {
                    foreach (var line in text.Skip(gr.Index).Take(gr.Count))
                        result.AppendLine(line);
                    continue;
                }

                try
                {
                    var indentationLength = Regex.Match(text[gr.Index], @"^\s*", RegexOptions.Multiline).Length;
                    var comment = XElement.Parse(
                        "<outer>{0}</outer>".Fmt(
                            text.Skip(gr.Index).Take(gr.Count)
                                .Select(line => Regex.Replace(line, @"^\s*/// ", "", RegexOptions.Multiline))
                                .JoinString(Environment.NewLine)
                        ),
                        LoadOptions.PreserveWhitespace
                    );

                    var wrapWidth = 126 - indentationLength;
                    var indentation = new string(' ', indentationLength) + "/// ";

                    // Special case: single <summary> tag that fits on a line
                    if (comment.Elements().Count() == 1 && comment.Elements().First().Name.LocalName == "summary" && !comment.Elements().First().Attributes().Any())
                    {
                        var summary = reformatComment(comment.Elements().First().Nodes(), false);
                        var wrapped = summary.WordWrap(wrapWidth - 4);
                        if (!wrapped.Skip(1).Any())
                        {
                            result.Append(indentation);
                            result.Append("<summary>");
                            result.Append(summary);
                            result.AppendLine("</summary>");
                            continue;
                        }
                    }

                    foreach (var line in reformatComment(comment.Nodes(), true).Trim().WordWrap(wrapWidth))
                        result.AppendLine(indentation + line);
                }
                catch (Exception e)
                {
                    result.AppendLine("The following comment is not valid: {0}".Fmt(e.Message, e.GetType().Name));
                    foreach (var line in text.Skip(gr.Index).Take(gr.Count))
                        result.AppendLine(line);
                }
            }

            return result.ToString();
        }

        private static string reformatComment(IEnumerable<XNode> nodes, bool topLevel, bool keepIndentation = false)
        {
            var sb = new StringBuilder();

            var isIn = Ut.Lambda((XNode node, string[] names) => names.Contains(((XElement) node).Name.LocalName));

            if (topLevel || nodes.All(n => (n is XText && string.IsNullOrWhiteSpace(((XText) n).Value)) || (n is XElement && !isIn(n, _inlineTags))) && nodes.Any(n => n is XElement && isIn(n, _blockLevelTags)))
            {
                // All tags are block-level or unknown ⇒ use block-level logic, which means:
                // • Discard all the whitespace between tags
                // • Put each opening tag on a new line and indent its contents
                var first = true;
                foreach (var elem in nodes.OfType<XElement>())
                {
                    if (!first)
                        sb.AppendLine();
                    first = false;
                    sb.Append(formatTag(elem, true, () =>
                    {
                        if (elem.Name.LocalName != "list")
                            return reformatComment(elem.Nodes(), false, keepIndentation: elem.Name.LocalName == "code").Indent(4);

                        // Handle <list> tags specially so that <item> and <description> don’t cause double-indentation
                        if (!elem.Nodes().All(n => n is XText || (n is XElement && ((XElement) n).Name.LocalName == "item")))
                            throw new InvalidOperationException("A “list” tag is not supposed to contain anything other than “item” tags.");
                        return elem.Nodes().OfType<XElement>().Select(e => "<item>{0}</item>".Fmt(reformatComment(e.Nodes(), false))).JoinString(Environment.NewLine).Indent(4);
                    }));
                }
                return sb.ToString();
            }
            else if (nodes.All(n => n is XText || (n is XElement && !isIn(n, _blockLevelTags))) && (!nodes.OfType<XElement>().Any() || nodes.OfType<XElement>().Any(n => isIn(n, _inlineTags))))
            {
                var first = true;
                string lastToAdd = null;

                // All nodes are text, inline-level or unknown ⇒ use inline-level logic, which means:
                // • Remove all single newlines (but keep double-newlines)
                // • Put all the nodes inline with the text
                foreach (var node in nodes)
                {
                    var text = node as XText;
                    var element = node as XElement;
                    if (text != null)
                    {
                        var value = !first ? text.Value : keepIndentation ? text.Value.TrimStart('\r', '\n') : text.Value.TrimStart();
                        sb.Append(lastToAdd);
                        if (keepIndentation)
                            lastToAdd = value.HtmlEscape(leaveSingleQuotesAlone: true, leaveDoubleQuotesAlone: true);
                        else
                            // Replace all “lone” newlines with spaces
                            lastToAdd = Regex.Replace(value, @"(?<!\n) *\r?\n *(?!\r?\n)", " ").HtmlEscape(leaveSingleQuotesAlone: true, leaveDoubleQuotesAlone: true);
                    }
                    else
                    {
                        sb.Append(lastToAdd);
                        lastToAdd = formatTag(element, false, () => reformatComment(element.Nodes(), false));
                    }
                    first = false;
                }
                sb.Append(lastToAdd.TrimEnd());

                var result = sb.ToString();
                var commonIndentation = Regex.Matches(result, @"^ *", RegexOptions.Multiline).Cast<Match>().Min(m => m.Length);
                return Regex.Replace(result, "^" + new string(' ', commonIndentation), "", RegexOptions.Multiline);
            }
            else
            {
                var firstBlock = (XElement) nodes.FirstOrDefault(n => n is XElement && isIn(n, _blockLevelTags));
                var firstInline = (XElement) nodes.FirstOrDefault(n => n is XElement && isIn(n, _inlineTags));
                if (firstBlock != null && firstInline != null)
                    throw new InvalidOperationException("“{0}” is block-level, but “{1}” is inline-level.".Fmt(firstBlock.Name.LocalName, firstInline.Name.LocalName));
                var firstUnknown = (XElement) nodes.FirstOrDefault(n => n is XElement && !isIn(n, _blockLevelTags) && !isIn(n, _inlineTags));
                throw new InvalidOperationException("I don’t know whether “{0}” is inline-level or block-level.".Fmt(firstUnknown.Name.LocalName));
            }
        }

        private static string formatTag(XElement elem, bool blockLevel, Func<string> inside)
        {
            if (!elem.Nodes().Any())
                // Self-closing tag
                return "<{0}{1}/>".Fmt(elem.Name.LocalName, elem.Attributes().Select(attr => @" {0}=""{1}""".Fmt(attr.Name.LocalName, attr.Value.HtmlEscape())).JoinString());

            return "<{0}{1}>{2}{3}</{0}>".Fmt(elem.Name.LocalName, elem.Attributes().Select(attr => @" {0}=""{1}""".Fmt(attr.Name.LocalName, attr.Value.HtmlEscape())).JoinString(), blockLevel ? Environment.NewLine : null, inside());
        }
    }
}
