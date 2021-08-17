﻿using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace RestartVisualStudio.Commands
{
    internal sealed class RestartNormalCommand
    {
        private readonly AsyncPackageBase _package;
        private static int CommandId => PackageIds.RestartNormalCommand;

        public RestartNormalCommand(AsyncPackageBase package, OleMenuCommandService commandService)
        {
            _package = package;

            var menuCommandId = new CommandID(PackageGuids.RestartVisualStudioCommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _package.RestartVisualStudio(elevated: false);
        }
        public static RestartNormalCommand Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackageBase package, OleMenuCommandService commandService)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            Instance = new RestartNormalCommand(package, commandService);
        }
    }
}