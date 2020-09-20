using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using GodotAddinVS.Debugging;
using Microsoft.VisualStudio;
using Mono.Debugging.Soft;
using Mono.Debugging.VisualStudio;
using Task = System.Threading.Tasks.Task;

namespace GodotAddinVS
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DebugCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 256;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("35470249-a450-49d4-9ce2-4486460267e9");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly GodotPackage package;

        private readonly IVsSolutionBuildManager vsSolutionBuildManager;

        /// <summary>
        /// Initializes a new instance of the <see cref="DebugCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DebugCommand(AsyncPackage package,
            OleMenuCommandService commandService,
            IVsSolutionBuildManager vsSolutionBuildManager)
        {
            this.package = package as GodotPackage ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            this.vsSolutionBuildManager = vsSolutionBuildManager ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DebugCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in DebugCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new DebugCommand(
                package,
                commandService,
                await package.GetServiceAsync(typeof(SVsSolutionBuildManager)) as IVsSolutionBuildManager);
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
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "DebugCommand";

            vsSolutionBuildManager.get_StartupProject(out var startupProject);

            package.GodotSolutionEventsListener.RegisterOpenGodotProject(startupProject);

            // https://stackoverflow.com/a/11061443
            // https://softwareproduction.eu/2012/09/28/convert-ivshierarchy-to-projectitem-or-project/
            startupProject.GetProperty(
                VSConstants.VSITEMID_ROOT,
                (int)__VSHPROPID.VSHPROPID_ExtObject,
                out object projObj);

            var proj = projObj as EnvDTE.Project;

            GodotDebuggableProjectCfg.DebugLaunchCore(proj);
        }
    }
}
