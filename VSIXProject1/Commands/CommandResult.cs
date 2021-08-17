namespace RestartVisualStudio.Commands
{
    public abstract class CommandResult
    {
        public string Message { get; set; }

        public CommandStatuses Status { get; set; }

        public bool Succeeded => this.Status == CommandStatuses.Success;
    }

    public class CancelledResult : CommandResult
    {
        public CancelledResult() => this.Status = CommandStatuses.Cancelled;
    }

    public class SuccessResult : CommandResult
    {
        public SuccessResult(string message = "")
        {
            this.Status = CommandStatuses.Success;
            this.Message = message;
        }
    }

    public class ProblemResult : CommandResult
    {
        public ProblemResult(string message)
        {
            this.Status = CommandStatuses.Problem;
            this.Message = message;
        }
    }
}
