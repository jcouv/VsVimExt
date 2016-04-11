/***************************************************************************
 
Copyright (c) Microsoft Corporation. All rights reserved.
THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

***************************************************************************/

using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsVimExtCommands
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    ///
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
    /// usually implemented by the package implementer.
    ///
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its 
    /// implementation of the IVsUIElementPane interface.
    /// </summary>
    [Guid("0bdb1e08-ed8b-47e8-91b2-e9bd814b4ebb")]
    public class FileOpenToolWindow : ToolWindowPane
    {
        private FileOpenControl control;

        /// <summary>
        /// Standard constructor for the tool window.
        /// </summary>
        public FileOpenToolWindow() :  
            base(null)
        {
            // Set the window title reading it from the resources.
            this.Caption = "Hello";

            // Set the image that will appear on the tab of the window frame
            // when docked with an other window
            // The resource ID correspond to the one defined in the resx file
            // while the Index is the offset in the bitmap strip. Each image in
            // the strip being 16x16.
            this.BitmapResourceID = 301;
            this.BitmapIndex = 0;

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
            // the object returned by the Content property.
            control = new FileOpenControl();
            base.Content = control; 
        }

        public override void OnToolWindowCreated()
        {
            base.OnToolWindowCreated();

            // Add our command handlers for toolbar buttons
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // TODO
            }

            // Add the toolbar to the window
            CreateToolBar();
        }

        private void CreateToolBar()
        {
            // Retrieve the shell UI object
            IVsUIShell4 shell4 = GetService(typeof(SVsUIShell)) as IVsUIShell4;
            if (shell4 != null)
            {
                // Create the toolbar tray
                IVsToolbarTrayHost host = null;
                if (ErrorHandler.Succeeded(shell4.CreateToolbarTray(this, out host)))
                {
                    // Add the toolbar as defined in vsct
                    host.AddToolbar(new Guid("00dcae55-4379-40a6-b152-3a38de753f29"), 0x2000);

                    IVsUIElement uiElement;
                    host.GetToolbarTray(out uiElement);

                    // Get the WPF element
                    object uiObject;
                    uiElement.GetUIObject(out uiObject);
                    IVsUIWpfElement wpfe = uiObject as IVsUIWpfElement;

                    // Retrieve and set the toolbar tray
                    object frameworkElement;
                    wpfe.GetFrameworkElement(out frameworkElement);
                    control.SetTray(frameworkElement as ToolBarTray);
                }
            }
        }
    }
}
