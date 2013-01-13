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
        private FontSupport[] _fontsSupported = new FontSupport[] { 
            new FontSupport { CommandName = "CourierNew", FontName = "Courier New", FontSize = 10 }, 
            new FontSupport { CommandName = "Candara", FontName = "Candara", FontSize = 11 }, 
            new FontSupport { CommandName = "SegoeUI", FontName = "Segoe UI", FontSize = 11, UseBold = false }, 
            new FontSupport { CommandName = "MaiandraGD", FontName = "Maiandra GD kun Eo", FontSize = 11 },
            new FontSupport { CommandName = "Georgia", FontName = "Georgia", FontSize = 11 }
        };

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
            var text = Regex.Split(source, @"\r?\n", RegexOptions.Singleline);
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
                var wrapWidth = 126 - indentation.Length;
                var wrap = Ut.Lambda((string str, int width) => str.Trim().WordWrap(width).ToArray());
                var chunk = XElement.Parse("<item>{0}</item>".Fmt(text.Subarray(gr.Index, gr.Count)
                    .Select(line => Regex.Replace(line, @"\s*/// ", "", RegexOptions.Multiline))
                    .JoinString(Environment.NewLine)), LoadOptions.PreserveWhitespace);
                var elements = chunk.Elements().ToArray();
                foreach (var elem in elements)
                    elem.Process();

                // Heuristic: If there is *only* a summary tag, and it fits on one line, make it compact
                if (elements.Length == 1 && elements[0].Name.LocalName == "summary" && wrap("<summary>" + elements[0].GetContent(), wrapWidth).Length == 1)
                    result.Add("{0}<summary>{1}</summary>".Fmt(indentation, elements[0].GetContent().Trim()));
                else
                {
                    var retStr = elements.Select(elem => "{0}{1}{2}</{3}>".Fmt(
                        elem.GetTag(),
                        Environment.NewLine,
                        wrap(elem.GetContent(), wrapWidth - 4).JoinString(Environment.NewLine).Indent(4),
                        elem.Name.LocalName
                    )).JoinString(Environment.NewLine);
                    result.Add(Regex.Replace(retStr, @"^", indentation, RegexOptions.Multiline));
                }
            }

            var resultStr = result.JoinString(Environment.NewLine);
            if (resultStr != source)
                startPoint.ReplaceText(endPoint, resultStr, (int) vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
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

        public static void Process(this XElement elem)
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
                        xElement.Process();
                    else if (xElement.Name.LocalName == "list")
                    {
                        foreach (var subElem in xElement.Elements("item"))
                            foreach (var subElem2 in subElem.Elements("description"))
                            {
                                subElem2.Process();
                                subElem2.FirstNode.AddBeforeSelf(new XText(Environment.NewLine + "    "));
                            }
                    }
                }
            }
        }
    }
}
