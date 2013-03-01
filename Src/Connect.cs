using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using EnvDTE;
using EnvDTE80;
using Extensibility;
using RT.Util;
using RT.Util.ExtensionMethods;

namespace RT.VisualStudio
{
    sealed class FontSupport
    {
        public string CommandName;
        public string FontName;
        public int FontSize;
        public bool UseBold = true;
    }

    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        private FontSupport[] _fontsSupported = Ut.NewArray(
            new FontSupport { CommandName = "CourierNew", FontName = "Courier New", FontSize = 10 },
            new FontSupport { CommandName = "Candara", FontName = "Candara", FontSize = 11 },
            new FontSupport { CommandName = "SegoeUI", FontName = "Segoe UI", FontSize = 11, UseBold = false },
            new FontSupport { CommandName = "MaiandraGD", FontName = "Maiandra GD kun Eo", FontSize = 11 },
            new FontSupport { CommandName = "Georgia", FontName = "Georgia", FontSize = 11 }
        );

        private string[] _thingsToBold = new[] { "Keyword", "User Types", "User Types(Value types)", "User Types(Interfaces)", "User Types(Delegates)", "User Types(Enums)", "User Types(Type parameters)" };
        private string[] _platformPriorities = new[] { "x86", "Any CPU", "AnyCPU" };

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private Dictionary<string, Action> _commands = new Dictionary<string, Action>();

        /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
        public Connect()
        {
        }

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2) application;
            _addInInstance = (AddIn) addInInst;

            if (connectMode == ext_ConnectMode.ext_cm_Startup)
            {
                CreateCommand("CloseAllToolWindows", "Close all Tool Windows", "Closes all tool windows.", () =>
                {
                    var windows = new List<Window>();
                    for (int i = 1; i <= _applicationObject.Windows.Count; i++)
                        if (_applicationObject.Windows.Item(i).Kind == "Tool")
                            windows.Add(_applicationObject.Windows.Item(i));
                    foreach (var window in windows)
                        window.Close();
                });

                foreach (var font in _fontsSupported)
                {
                    CreateCommand("ChangeFontTo" + font.CommandName, "Change Font to " + font.FontName, string.Format("Changes the text editor font to {0}.", font.FontName), () =>
                    {
                        foreach (Property prop in _applicationObject.Properties["FontsAndColors", "TextEditor"])
                        {
                            if (prop.Name == "FontFamily")
                                prop.Value = font.FontName;
                            else if (prop.Name == "FontSize")
                                prop.Value = font.FontSize;
                            else if (prop.Name == "FontsAndColorsItems")
                            {
                                var o = (FontsAndColorsItems) prop.Object;
                                foreach (ColorableItems obj in o)
                                    if (_thingsToBold.Contains(obj.Name))
                                    {
                                        obj.Bold = font.UseBold;
                                        obj.Background = 0x2000000;
                                    }
                            }
                        }
                    });
                }

                CreateCommand("ReformatXmlComments", "Reformat XML Comments", "Automatically word-wraps and reformats XML comments to conform to the RT comment style.", reformatComments);

                //CreateCommand("CleanUpSolutionConfigurations", "Clean Up Solution Configurations", "Removes all solution configurations except for “Debug” and “Release”.", () =>
                //{
                //    var configurations = _applicationObject.Solution.SolutionBuild.SolutionConfigurations;

                //    // Check if Debug|x86 and Release|x86 are present
                //    var debug = configurations.Cast<SolutionConfiguration>().Where(sc => sc.Name == "Debug").ToArray();
                //    var release = configurations.Cast<SolutionConfiguration>().Where(sc => sc.Name == "Release").ToArray();

                //    if (debug.Length == 0 || release.Length == 0)
                //    {
                //        MessageBox.Show("There are no two solution configurations called “Debug” and “Release”, respectively.");
                //        return;
                //    }

                //    var pairs = debug
                //        .Select(d =>
                //        {
                //            var platform = ((dynamic) d).PlatformName;
                //            return new
                //            {
                //                Debug = d,
                //                Release = release.FirstOrDefault(r => ((dynamic) r).PlatformName == platform),
                //                Platform = platform,
                //                Priority = Array.IndexOf(_platformPriorities, platform)
                //            };
                //        })
                //        .Where(inf => inf.Release != null)
                //        .OrderByDescending(inf => inf.Priority)
                //        .ToArray();

                //    if (pairs.Length == 0)
                //    {
                //        MessageBox.Show("There are no two solution configurations called “Debug” and “Release” with the same platform.");
                //        return;
                //    }

                //    if (pairs.Length > 1 && pairs[0].Priority == pairs[1].Priority)
                //    {
                //        MessageBox.Show(string.Format("There are two solution configurations with the same platform priority ({0} and {1}).", pairs[0].Platform, pairs[1].Platform));
                //        return;
                //    }

                //    var preferred = pairs[0];

                //    // Delete all other configurations
                //    while (true)
                //    {
                //        try
                //        {
                //            var unwantedConfig = configurations.Cast<SolutionConfiguration>().FirstOrDefault(sc => sc != preferred.Debug && sc != preferred.Release);
                //            if (unwantedConfig == null)
                //                break;
                //            ((SolutionConfiguration2) unwantedConfig).Delete();
                //        }
                //        catch (Exception e)
                //        {
                //            System.Diagnostics.Debugger.Break();
                //        }
                //    }

                //    System.Diagnostics.Debugger.Break();
                //});
            }
        }

        private void reformatComments()
        {
            var doc = (TextDocument) _applicationObject.ActiveDocument.Object();
            var startPoint = doc.CreateEditPoint();
            startPoint.StartOfDocument();
            var endPoint = doc.CreateEditPoint();
            endPoint.EndOfDocument();
            var source = startPoint.GetText(endPoint);
            var resultStr = reformatComments(source);
            if (resultStr != source)
                startPoint.ReplaceText(endPoint, resultStr, (int) vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
        }

        private static string[] _inlineTags = new[] { "see", "paramref", "typeparamref", "c" };
        private static string[] _blockLevelTags = new[] { "code", "para", "list", "description" };

        private static string reformatComments(string source)
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
                    sb.AppendLine("<{0}{1}>".Fmt(elem.Name.LocalName, elem.Attributes().Select(attr => @" {0}=""{1}""".Fmt(attr.Name.LocalName, attr.Value.HtmlEscape())).JoinString()));

                    string inner;
                    if (elem.Name.LocalName == "list")
                    {
                        // Handle <list> tags specially so that <item> and <description> don’t cause double-indentation
                        if (!elem.Nodes().All(n => n is XText || (n is XElement && ((XElement) n).Name.LocalName == "item")))
                            throw new InvalidOperationException("A “list” tag is not supposed to contain anything other than “item” tags.");
                        inner = elem.Nodes().OfType<XElement>().Select(e => "<item>{0}</item>".Fmt(reformatComment(e.Nodes(), false))).JoinString(Environment.NewLine);
                    }
                    else
                        inner = reformatComment(elem.Nodes(), false, keepIndentation: elem.Name.LocalName == "code");

                    sb.Append(inner.Indent(4));
                    sb.AppendFormat("</{0}>", elem.Name.LocalName);
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
                            lastToAdd = value;
                        else
                            // Replace all “lone” newlines with spaces
                            lastToAdd = Regex.Replace(value, @"(?<!\n) *\r?\n *(?!\r?\n)", " ");
                    }
                    else
                    {
                        sb.Append(lastToAdd);
                        lastToAdd = element.ToString();
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

        private void CreateCommand(string commandName, string readableCommandName, string commandDescription, Action action)
        {
            try
            {
                object[] blah = { };
                ((Commands2) _applicationObject.Commands).AddNamedCommand2(_addInInstance, commandName, readableCommandName, commandDescription, true, Type.Missing, ref blah, 3, 3, vsCommandControlType.vsCommandControlTypeButton);
            }
            catch (ArgumentException)
            {
                //If we are here, then the exception is probably because a command with that name
                //  already exists. If so there is no need to recreate the command and we can 
                //  safely ignore the exception.
            }

            _commands[typeof(Connect).FullName + "." + commandName] = action;
        }

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param term='commandName'>The name of the command to determine state for.</param>
        /// <param term='neededText'>Text that is needed for the command.</param>
        /// <param term='status'>The state of the command in the user interface.</param>
        /// <param term='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (_commands.ContainsKey(commandName))
                {
                    status = (vsCommandStatus) vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
            }
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param term='commandName'>The name of the command to execute.</param>
        /// <param term='executeOption'>Describes how the command should be run.</param>
        /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param term='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (_commands.ContainsKey(commandName))
                {
                    _commands[commandName]();
                    handled = true;
                }
            }
        }
    }
}
