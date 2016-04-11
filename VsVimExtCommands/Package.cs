//------------------------------------------------------------------------------
// <copyright file="MyCommandPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace VsVimExtCommands
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(CommandPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideToolWindow(typeof(FileOpenToolWindow), MultiInstances=true)]
    public sealed class CommandPackage : Package
    {
        /// <summary>
        /// MyCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "653fd200-5661-4ab2-a634-f6c97a6af1ca";

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// </summary>
        public CommandPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// This function is called when the user clicks the menu item that shows the 
        /// tool window. See the Initialize method to see how the menu item is associated to 
        /// this function using the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void ShowToolWindow(object sender, EventArgs e)
        {
            // For a multi-instance ToolWindow, find an unused ID
            int id = FindUnusedToolWindowId(typeof(FileOpenToolWindow));

            // Create the window with the unused ID.
            var window = CreateToolWindow(typeof(FileOpenToolWindow), id) as FileOpenToolWindow;
            if ((null == window) || (null == window.Frame))
            {
                //throw new NotSupportedException(Resources.CanNotCreateWindow);
            }

            // Display the window.
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        /// <summary>
        /// Find an unused ID for a new multi instance toolwindow.
        /// </summary>
        /// <param name="toolWindowType">The type of the toolwindow</param>
        /// <returns>An unused ID</returns>
        private int FindUnusedToolWindowId(Type toolWindowType)
        {
            for (int id = 0; ; ++id)
            {
                ToolWindowPane window = FindToolWindow(toolWindowType, id, false);
                if (window == null)
                {
                    return id;
                }
            }
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            var service = GetService(typeof(SVsUIShell)) as IVsUIShell;

            IEnumWindowFrames windows;
            int i = service.GetDocumentWindowEnum(out windows);

            OleMenuCommandService commandService = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID3 = new CommandID(Command.CommandSet, Command.CommandId3);
                var menuItem3 = new MenuCommand(ShowToolWindow, menuCommandID3);
                commandService.AddCommand(menuItem3);
            }
                Command.Initialize(this);
            base.Initialize();
        }

        #endregion
    }
}
