/// <summary>
/// Email This DocumentCodeSpanCommand Class
/// </summary>
namespace EmailThis
{
    using System;
    using System.ComponentModel.Design;
    using System.IO;
    using EnvDTE;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.TextManager.Interop;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler.
    /// </summary>
    internal sealed class DocumentCodeSpanCommand
    {
        /// <summary>
        /// Defines the <see cref="TextViewPosition" />.
        /// </summary>
        public struct TextViewPosition
        {
            /// <summary>
            /// Defines the _column.
            /// </summary>
            private readonly int _column;

            /// <summary>
            /// Defines the _line.
            /// </summary>
            private readonly int _line;

            /// <summary>
            /// Initializes a new instance of the <see cref=""/> class.
            /// </summary>
            /// <param name="line">The line<see cref="int"/>.</param>
            /// <param name="column">The column<see cref="int"/>.</param>
            public TextViewPosition(int line, int column)
            {
                _line = line;
                _column = column;
            }

            /// <summary>
            /// Gets the Line.
            /// </summary>
            public int Line
            {
                get { return _line; }
            }

            /// <summary>
            /// Gets the Column.
            /// </summary>
            public int Column
            {
                get { return _column; }
            }

            /// <summary>
            /// Implements the operator &lt;.
            /// </summary>
            /// <param name="a">a.</param>
            /// <param name="b">The b.</param>
            /// <returns>
            /// The result of the operator.
            /// </returns>
            public static bool operator <(TextViewPosition a, TextViewPosition b)
            {
                if (a.Line < b.Line)
                {
                    return true;
                }
                else if (a.Line == b.Line)
                {
                    return a.Column < b.Column;
                }
                else
                {
                    return false;
                }
            }

            /// <summary>
            /// Implements the operator &gt;.
            /// </summary>
            /// <param name="a">a.</param>
            /// <param name="b">The b.</param>
            /// <returns>
            /// The result of the operator.
            /// </returns>
            public static bool operator >(TextViewPosition a, TextViewPosition b)
            {
                if (a.Line > b.Line)
                {
                    return true;
                }
                else if (a.Line == b.Line)
                {
                    return a.Column > b.Column;
                }
                else
                {
                    return false;
                }
            }

            /// <summary>
            /// The Min.
            /// </summary>
            /// <param name="a">The a<see cref="TextViewPosition"/>.</param>
            /// <param name="b">The b<see cref="TextViewPosition"/>.</param>
            /// <returns>The <see cref="TextViewPosition"/>.</returns>
            public static TextViewPosition Min(TextViewPosition a, TextViewPosition b)
            {
                return a > b ? b : a;
            }

            /// <summary>
            /// The Max.
            /// </summary>
            /// <param name="a">The a<see cref="TextViewPosition"/>.</param>
            /// <param name="b">The b<see cref="TextViewPosition"/>.</param>
            /// <returns>The <see cref="TextViewPosition"/>.</returns>
            public static TextViewPosition Max(TextViewPosition a, TextViewPosition b)
            {
                return a > b ? a : b;
            }
        }

        /// <summary>
        /// Defines the <see cref="TextViewSelection" />.
        /// </summary>
        public struct TextViewSelection
        {
            /// <summary>
            /// Gets or sets the StartPosition.
            /// </summary>
            public TextViewPosition StartPosition { get; set; }

            /// <summary>
            /// Gets or sets the EndPosition.
            /// </summary>
            public TextViewPosition EndPosition { get; set; }

            /// <summary>
            /// Gets or sets the Text.
            /// </summary>
            public string Text { get; set; }

            /// <summary>
            /// Initializes a new instance of the <see cref=""/> class.
            /// </summary>
            /// <param name="a">The a<see cref="TextViewPosition"/>.</param>
            /// <param name="b">The b<see cref="TextViewPosition"/>.</param>
            /// <param name="text">The text<see cref="string"/>.</param>
            public TextViewSelection(TextViewPosition a, TextViewPosition b, string text)
            {
                StartPosition = TextViewPosition.Min(a, b);
                EndPosition = TextViewPosition.Max(a, b);
                Text = text;
            }
        }

        /// <summary>
        /// Command ID..
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID)..
        /// </summary>
        public static readonly Guid CommandSet = new Guid("84ae883f-2fd6-41ee-a0b9-939f2f861916");

        /// <summary>
        /// VS Package that provides this command, not null..
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Prevents a default instance of the <see cref="DocumentCodeSpanCommand"/> class from being created.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DocumentCodeSpanCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            ThreadHelper.ThrowIfNotOnUIThread();

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command..
        /// </summary>
        public static DocumentCodeSpanCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package..
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>The <see cref="Task"/>.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in DocumentCodeSpan's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new DocumentCodeSpanCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            TextViewSelection selection = GetSelection(ServiceProvider);
            var activeDocument = GetActiveFile(ServiceProvider);
            SendEmailTo(activeDocument, selection);
        }

        /// <summary>
        /// The GetSelection.
        /// </summary>
        /// <param name="serviceProvider">The serviceProvider<see cref="IServiceProvider" />.</param>
        /// <returns>
        /// The <see cref="TextViewSelection" />.
        /// </returns>
        private TextViewSelection GetSelection(IServiceProvider serviceProvider)
        {
            var service = serviceProvider.GetService(typeof(SVsTextManager));
            var textManager = service as IVsTextManager2;
            IVsTextView view;
            textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out view);
            view.GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);//end could be before beginning
            var start = new TextViewPosition(startLine, startColumn);
            var end = new TextViewPosition(endLine, endColumn);

            view.GetSelectedText(out string selectedText);

            TextViewSelection selection = new TextViewSelection(start, end, selectedText);
            return selection;
        }

        /// <summary>
        /// The GetActiveFile.
        /// </summary>
        /// <param name="serviceProvider">The serviceProvider<see cref="IServiceProvider" />.</param>
        /// <returns>
        /// The <see cref="Document" />.
        /// </returns>
        private Document GetActiveFile(IServiceProvider serviceProvider)
        {
            EnvDTE80.DTE2 applicationObject = serviceProvider.GetService(typeof(DTE)) as EnvDTE80.DTE2;
            return applicationObject?.ActiveDocument;
        }

        /// <summary>
        /// The SendEmailTo.
        /// </summary>
        /// <param name="activeDocument">The activeDocument<see cref="Document" />.</param>
        /// <param name="selection">The selection<see cref="TextViewSelection" />.</param>
        private void SendEmailTo(Document activeDocument, TextViewSelection selection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var selectionText = string.Empty;
            if (selection.Text.Length == 0)
            {
                selectionText = File.ReadAllText(activeDocument.FullName);
            }
            else
            {
                selectionText = selection.Text;
            }

            if (selectionText?.Length > 18000)
            {
                selectionText = "-- Content is trimmed due to large size -- \n\n" + selectionText.Substring(0, 18000);
            }

            string command =
                string.Format("mailto:?subject={0}&body={1}",
                    Uri.EscapeDataString("Check this Code Snippet : " + activeDocument.Name),
                    Uri.EscapeDataString(selectionText));
            System.Diagnostics.Process.Start(command);
        }
    }
}
