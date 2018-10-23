#if DEBUG

namespace Routines.Excel.EventsIndexing.Tests
{
    public class BasingEventsLauncher : FullDbLauncher
    {
        public override string FolderOut
        {
            get { return "out"; }
        }
    }
}
#endif