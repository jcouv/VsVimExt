using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel.Design;

namespace VsVimExtCommands
{
    internal sealed class Command
    {
        public const int CommandId = 0x0100;
        public const int CommandId2 = 0x0101;
        public const int CommandId3 = 0x0102;

        public static readonly Guid CommandSet = new Guid("e8bbf02f-7fa8-46e3-af8e-92d39ab79665");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.NextErrorInView, menuCommandID);
                commandService.AddCommand(menuItem);

                var menuCommandID2 = new CommandID(CommandSet, CommandId2);
                var menuItem2 = new MenuCommand(this.PreviousErrorInView, menuCommandID2);
                commandService.AddCommand(menuItem2);

                var menuCommandID3 = new CommandID(CommandSet, CommandId3);
                var menuItem3 = new MenuCommand(this.MoveTabLeft, menuCommandID3);
                commandService.AddCommand(menuItem3);
            }
        }

        public static Command Instance
        {
            get;
            private set;
        }

        private IServiceProvider ServiceProvider => this.package;

        public static void Initialize(Package package)
        {
            Instance = new Command(package);
        }

        // TODO: the commands only appear to load when I use the menu, not the VsVim keybindings (ge and gE).
        // TODO: move tab to the left/right
        // TODO: file explorer with Vim-style bindings
        // TODO: list derived types
        // TODO: next/previous error in document (not just in view)

        private void MoveTabLeft(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void PreviousErrorInView(object sender, EventArgs e)
        {
            IVsTextView textView = GetTextView();
            var layer = GetAdornmentLayer(textView);

            var cursorPosition = GetCursorPosition(textView);
            var maxPreviousPosition = 0;

            foreach (var element in layer.Elements)
            {
                var adornment = element.Adornment.GetType().Name;

                var elementPosition = element.VisualSpan.Value.Start.Position;
                if (elementPosition < cursorPosition && elementPosition > maxPreviousPosition)
                {
                    maxPreviousPosition = elementPosition;
                }
            }

            if (maxPreviousPosition != 0)
            {
                SetCursorPosition(maxPreviousPosition, textView);
            }
        }

        private void NextErrorInView(object sender, EventArgs e)
        {
            IVsTextView textView = GetTextView();
            var layer = GetAdornmentLayer(textView);

            var cursorPosition = GetCursorPosition(textView);

            int minNextPosition = int.MaxValue;
            foreach (var element in layer.Elements)
            {
                var elementPosition = element.VisualSpan.Value.Start.Position;
                if (elementPosition > cursorPosition && elementPosition < minNextPosition)
                {
                    minNextPosition = elementPosition;
                }
            }

            if (minNextPosition != int.MaxValue)
            {
                SetCursorPosition(minNextPosition, textView);
            }
        }

        private static IVsTextView GetTextView()
        {
            int err;
            var textManager = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));

            IVsTextView textView;
            err = textManager.GetActiveView(1, null, out textView);

            return textView;
        }

        private void SetCursorPosition(int position, IVsTextView textView)
        {
            int err, line, column;
            err = textView.GetLineAndColumn(position, out line, out column);
            err = textView.SetCaretPos(line, column);
        }

        private int GetCursorPosition(IVsTextView textView)
        {
            int err;
            int line, column;
            err = textView.GetCaretPos(out line, out column);

            int virtualSpaces;
            int cursorPosition;
            textView.GetNearestPosition(line, column, out cursorPosition, out virtualSpaces);

            return cursorPosition;
        }

        private IAdornmentLayer GetAdornmentLayer(IVsTextView textView)
        {
            IVsUserData userData = textView as IVsUserData;

            object holder;
            Guid guidViewHost = DefGuidList.guidIWpfTextViewHost;
            userData.GetData(ref guidViewHost, out holder);

            IWpfTextViewHost viewHost = (IWpfTextViewHost)holder;
            return viewHost.TextView.GetAdornmentLayer(PredefinedAdornmentLayers.Squiggle);
        }

        //Microsoft.VisualStudio.Shell.servi
        //IVsUIShell shell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
        //IEnumWindowFrames windows;
        //int err = shell.GetDocumentWindowEnum(out windows);
        //int i = 0;

        //IVsWindowFrame[] frames = new IVsWindowFrame[1];
        //uint numFrames;
        //while ((windows.Next(1, frames, out numFrames) == VSConstants.S_OK) && (numFrames == 1))
        //{
        //    if (frames[0].IsVisible() == VSConstants.S_OK)
        //    {
        //        object t;
        //        frames[0].GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out t);


        //        object docView;
        //        frames[0].GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out docView);
        //        IVsCodeWindow codeWindow = docView as IVsCodeWindow;

        //        IVsTextLines lines;
        //        err = codeWindow.GetBuffer(out lines);
        //    }
        //}

        //IWpfTextView wpfTextView = (IWpfTextView)textView;


        //var adornmentLayer = wpfTextView.GetAdornmentLayer(PredefinedAdornmentLayers.Squiggle);

        //IVsEditorAdaptersFactoryService editorAdaptersFactoryService = Package.GetGlobalService(typeof(IVsEditorAdaptersFactoryService)) as IVsEditorAdaptersFactoryService;

        //DTE dte = (DTE)Package.GetGlobalService(typeof(DTE));

        //ServiceProvider sp = new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)dte);
        //if (sp != null)
        //{
        //    IVsActivityLog log = sp.GetService(typeof(SVsActivityLog)) as IVsActivityLog;
        //    if (log != null)
        //    {
        //        System.Windows.Forms.MessageBox.Show("Found the activity log service.");
        //    }
        //}

        //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
        //string title = $"MyCommand {i}";

        //// Show a message box to prove we were here
        //VsShellUtilities.ShowMessageBox(
        //    this.ServiceProvider,
        //    message,
        //    title,
        //    OLEMSGICON.OLEMSGICON_INFO,
        //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
        //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);



        //Microsoft.VisualStudio.TextManager.Interop.IVsCodeWindow window = pane as Microsoft.VisualStudio.TextManager.Interop.IVsCodeWindow;

        //if (window != null)
        //{

        //    Microsoft.VisualStudio.TextManager.Interop.IVsTextLines lines;
        //    window.GetBuffer(out lines);
        //    lines.CreateLineMarker((int)Microsoft.VisualStudio.TextManager.Interop.MARKERTYPE.MARKER_OTHER_ERROR, 0, 0, 1, 0, null, null);
        //    lines.FindMarkerByLineIndex((int)Microsoft.VisualStudio.TextManager.Interop.MARKERTYPE.MARKER_OTHER_ERROR, 0, 0, 0, out marker);
        //}

    }
}
