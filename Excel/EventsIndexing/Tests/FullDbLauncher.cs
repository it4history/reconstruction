#if DEBUG

namespace Routines.Excel.EventsIndexing.Tests
{
    public class FullDbLauncher : LauncherBase
    {
        public override string FileNameIn
        {
            get { return "00_����_2018_10_14.xls"; }
        }
        public override string FolderOut
        {
            get { return "out"; }
        }
    }
}
#endif