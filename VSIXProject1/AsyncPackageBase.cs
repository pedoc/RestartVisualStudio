using System;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using RestartVisualStudio.Commands;

namespace RestartVisualStudio
{
    public class AsyncPackageBase : AsyncPackage
    {
        public CommandResult RestartVisualStudio(
            bool confirm = true,
            bool saveFiles = true,
            bool elevated = false)
        {
            try
            {
                if (confirm && !this.GetRestartConfirmation(elevated))
                    return (CommandResult)new CancelledResult();
                if (saveFiles && !this.SaveAllFiles().Succeeded)
                    return (CommandResult)new ProblemResult("Unable to save open files");
                return elevated ? this.RestartElevated() : this.RestartNormal();
            }
            catch (Exception ex)
            {
                return (CommandResult)new ProblemResult(ex.Message);
            }
        }

        private DTE2 _dte;
        protected DTE2 Dte => _dte ?? (_dte = GetGlobalService(typeof(DTE2)) as DTE2);
        private IVsShell4 _vsShell;
        protected IVsShell4 VsShell
        {
            get
            {
                IVsShell4 vsShell = this._vsShell;
                if (vsShell != null)
                    return vsShell;
                IVsShell4 service;
                this._vsShell = service = this.GetService<SVsShell, IVsShell4>();
                return service;
            }
        }

        private CommandResult RestartNormal(string problem = "") => this.Restart(0U, problem);

        private CommandResult RestartElevated(string problem = "") => this.Restart(1U, problem);

        private CommandResult Restart(uint mode, string problem = "")
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return ErrorHandler.Failed(this.VsShell.Restart(mode)) ? (CommandResult)new ProblemResult(problem) : (CommandResult)new SuccessResult();
            }
            catch (Exception ex)
            {
                return (CommandResult)new ProblemResult(problem ?? ex.Message);
            }
        }

        public CommandResult SaveAllFiles(string success = "", string problem = "")
        {
            try
            {
                return this.ExecuteCommand("File.SaveAll", success: success, problem: problem);
            }
            catch (Exception ex)
            {
                return (CommandResult)new ProblemResult(problem ?? ex.Message);
            }
        }

        public CommandResult ExecuteCommand(
            string command,
            string args = "",
            string success = null,
            string problem = null)
        {
            try
            {
                this.Dte?.ExecuteCommand(command, args);
                return (CommandResult)new SuccessResult(success);
            }
            catch (Exception ex)
            {
                return (CommandResult)new ProblemResult(problem ?? ex.Message);
            }
        }

        private bool GetRestartConfirmation(bool elevated) => this.DisplayQuestion("Question", (
            $"You're about to restart Visual Studio{(elevated ? (object)" As Administrator" : (object)"")}. " + "Any open files will automatically be saved for you first though."));


        public bool DisplayQuestion(string title, string message)
        {
            var localTitle = L(title);
            var localMessage = L(message);

            var result = DisplayMessage(this, localTitle, localMessage, OLEMSGBUTTON.OLEMSGBUTTON_YESNO, OLEMSGICON.OLEMSGICON_QUERY);
            return result == 6;
        }

        private static string L(string value)
        {
            var result = Properties.Resources.ResourceManager.GetString(value);
            if (result == null) return value;
            return result;
        }

        public static int DisplayMessage(
            AsyncPackageBase package,
            string title = null,
            string message = "",
            OLEMSGBUTTON button = OLEMSGBUTTON.OLEMSGBUTTON_OK,
            OLEMSGICON icon = OLEMSGICON.OLEMSGICON_INFO)
        {
            return VsShellUtilities.ShowMessageBox(
                 package,
                 message,
                 title,
                 icon,
                 button,
                 OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
