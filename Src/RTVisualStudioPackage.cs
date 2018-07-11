using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using RT.Util;

namespace RT.VisualStudio
{
    /// <summary>
    ///     This is the class that implements the package exposed by this assembly.</summary>
    /// <remarks>
    ///     <para>
    ///         The minimum requirement for a class to be considered a valid package for Visual Studio is to implement the
    ///         IVsPackage interface and register itself with the shell. This package uses the helper classes defined inside
    ///         the Managed Package Framework (MPF) to do it: it derives from the Package class that provides the
    ///         implementation of the IVsPackage interface and uses the registration attributes defined in the framework to
    ///         register itself and its components with the shell. These attributes tell the pkgdef creation utility what data
    ///         to put into .pkgdef file.</para>
    ///     <para>
    ///         To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage"
    ///         ...&gt; in .vsixmanifest file.</para></remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid("1c3f0810-9400-4763-8903-07cc8d9281b7")]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class RTVisualStudioPackage : Package
    {
        private FontSupport[] _fontsSupported = Ut.NewArray(
            new FontSupport { CommandName = "CourierNew", FontName = "Courier New", FontSize = 10 },
            new FontSupport { CommandName = "Cambria", FontName = "Cambria", FontSize = 12, UseBold = false },
            new FontSupport { CommandName = "Georgia", FontName = "Georgia", FontSize = 11, UseBold = false }
        );
        private string[] _thingsToBold = new[] { "Keyword", "User Types", "User Types(Value types)", "User Types(Interfaces)", "User Types(Delegates)", "User Types(Enums)", "User Types(Type parameters)" };

        public RTVisualStudioPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        ///     Initialization of the package; this method is called right after the package is sited, so this is the place
        ///     where you can put all the initialization code that rely on services provided by VisualStudio.</summary>
        protected override void Initialize()
        {
            var commandService = GetService(typeof(IMenuCommandService)) as IMenuCommandService;
            if (commandService != null)
            {
                // CloseAllToolWindows
                commandService.AddCommand(new MenuCommand(
                    command: new CommandID(
                        menuGroup: new Guid("37723bb9-70e2-4825-aa61-536e6401c65c"),
                        commandID: 0x0100),
                    handler: (_, __) =>
                    {
                        dynamic dte = GetService(typeof(SDTE));
                        var windows = new List<Window>();
                        for (int i = 1; i <= dte.Windows.Count; i++)
                        {
                            if (dte.Windows.Item(i) != null && dte.Windows.Item(i).Kind == "Tool")
                                windows.Add(dte.Windows.Item(i));
                        }
                        foreach (var window in windows)
                            window.Close();
                    }));

                // ReformatComments
                commandService.AddCommand(new MenuCommand(
                    command: new CommandID(
                        menuGroup: new Guid("37723bb9-70e2-4825-aa61-536e6401c65c"),
                        commandID: 0x0101),
                    handler: (_, __) =>
                    {
                        dynamic dte = GetService(typeof(SDTE));
                        var doc = (TextDocument) dte.ActiveDocument.Object();
                        var startPoint = doc.CreateEditPoint();
                        startPoint.StartOfDocument();
                        var endPoint = doc.CreateEditPoint();
                        endPoint.EndOfDocument();
                        var source = startPoint.GetText(endPoint);
                        var resultStr = CommentFormatter.ReformatComments(source);
                        if (resultStr != source)
                            startPoint.ReplaceText(endPoint, resultStr, (int) vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
                    }));

                // Font commands
                for (int i = 0; i < _fontsSupported.Length; i++)
                {
                    var font = _fontsSupported[i];
                    commandService.AddCommand(new MenuCommand(
                        command: new CommandID(
                            menuGroup: new Guid("37723bb9-70e2-4825-aa61-536e6401c65c"),
                            commandID: 0x0200 + i),
                        handler: (_, __) =>
                        {
                            dynamic dte = GetService(typeof(SDTE));
                            foreach (Property prop in dte.Properties["FontsAndColors", "TextEditor"])
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
                        }));
                }
            }

            base.Initialize();
        }
    }
}
