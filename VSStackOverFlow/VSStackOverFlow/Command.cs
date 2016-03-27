//------------------------------------------------------------------------------
// <copyright file="Command.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;


namespace VSStackOverFlow
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("256a8dcb-8d41-4e77-a00f-0f05e9e63d35");

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
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
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
        public static void Initialize(Package package)
        {
            Instance = new Command(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {

            //Get the output pane
            var outputWindow = (ErrorList)this.ServiceProvider.GetService(typeof(SVsErrorList));
            var listy = outputWindow.ErrorItems;

            //make sure we don't go out of range
            if(listy.Count == 0)
            {
                return;
            }

            //Get the first one.
            var item1 = listy.Item(1);

            //Parse the file path to determine the project type.
            var fileUri = new Uri(item1.FileName);

            //Check the file extension.
            string fileNameWithExtension = fileUri.Segments[fileUri.Segments.Length - 1];
            string fileExtension = fileNameWithExtension.Substring(fileNameWithExtension.IndexOf("."));

            string languageOfChoice = string.Empty;

            if (fileExtension == ".cs")
            {
                languageOfChoice = "C%23"; // C# on query string will appear as C%23
            }

            //get the description of the error.
            string errorDescription = item1.Description;

            //build the query string.
            string prefixUrl = "http://stackoverflow.com/search?q=";
            string url = prefixUrl + languageOfChoice + "+" + errorDescription;


            //Create the browser
            IVsWindowFrame browserFrame;

            //Get a service for browsing session
            var browserService = Package.GetGlobalService(typeof(IVsWebBrowsingService)) as IVsWebBrowsingService;

            //Navigate and output it to a frame.
            //0 = force to create the tab if it isn't exist, and use existing one if there is one.
            browserService.Navigate(url, 0, out browserFrame);
            //string message = "Hello World!";
            //string title = "Send from my Surface to your face";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
