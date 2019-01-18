using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using DevExpress.Xpo;
using Logy.Entities.Engine;
using Logy.Entities.Import;
using Logy.Entities.Localization;
using Logy.Entities.Model;
using Logy.MwAgent.DotNetWikiBot;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Logy.Api.Mw.Excel
{
    public class ExcelImporter : CategoryImporter
    {
        private ExcelManager _excelManager;

        public ExcelImporter(Session session, Guid? wikiId = null) : base(session, wikiId)
        {
        }

        public override string Url
        {
            get { return "Category:Excel import"; }
        }

        public override bool IsGroup { get { return true; } }

        public override EntityType EntityType
        {
            get { return EntityType.None; }
        }

        public override Speed GetPagesSpeed
        {
            // to allow during Excel_ call of GetPages() and excel file to load
            get { return Speed.Fast; }
        }

        public string Dir
        {
            get
            {
                return Path.Combine(
                    Environment.CurrentDirectory,
                    WikiSite.ShortPath.TrimStart('/'));
            }
        }

        public string Filepath
        {
            get
            {
                return Path.Combine(Dir, Template.Url);
            }
        }

        public ExcelManager ExcelManager
        {
            get
            {
                if (_excelManager == null)
                {
                    _excelManager = new ExcelManager(Filepath);
                    _excelManager.Read();
                }
                return _excelManager;
            }
        }

        public override IList GetPages(ImportStatus status = null)
        {
            string result = null;
            if (WikiSite == null)
                return null;

            if (!Directory.Exists(Dir))
                Directory.CreateDirectory(Dir);
            if (!File.Exists(Filepath))
            {
                result = new Page(WikiSite, "File:" + Template.Url).DownloadImage(Filepath);
                if (status != null)
                {
                    status.Error = result;
                }
            }

            if (result == null)
            {
                foreach (HSSFRow row in ExcelManager.Records)
                {
                    var title = GetTitle(row);
                    var yearsInTitle = ExcelManager.GetYears(title);
                    var yearInColumn = GetYear(row);
                    if (yearsInTitle != null
                        && !string.IsNullOrEmpty(yearInColumn)
                        && !yearsInTitle.Contains(yearInColumn))
                    {
                        if (status != null)
                        {
                            status.Warning = string.Format(
                                "row #{0} has dates conflict: in description {1} but in year {2}",
                                row.RowNum,
                                yearsInTitle,
                                yearInColumn);
                        }
                    }
                }
            }
            if (result == null)
            {
                if (status != null)
                {
                    status.Result = ExcelManager.Columns;
                }
                return ExcelManager.Records;
            }
            return null;
        }

        public override void Import(object page, Job job)
        {
            var row = (HSSFRow)page;
            var man = new EventsFromExcelManager(this, row, job);
            if (man.WikiTemplateParametersModified.Count > 0 || man.WikiPageCreated)
            {
                if (man.WikiPageCreated)
                    ObjectAdded(man.Title);
                if (man.WikiTemplateParametersModified.Count > 0)
                    ObjectUpdated(man.Title);

                man.WikiPage.Save(
                    "[[" + man.ThisExcelFile + "]]; columns: " +
                    man.WikiTemplateParametersModified.Aggregate((c, s) => c + ", " + s),
                    true);
            }
        }

        internal Page NewPage(ref string title)
        {
            while (Encoding.UTF8.GetByteCount(title) > 255)
            {
                title = title.Remove(title.Length - 1);
            }
            var page = new Page(WikiSite, title);
            page.LoadTextOnly();
            return page;
        }

        internal string GetTitle(HSSFRow row)
        {
            var descriptionColumn = JsonManager.GetJsonTranslation(ExcelFileColumns.Description);
            var titleValue = ExcelManager.GetValue(
                row,
                descriptionColumn);
            return ExcelManager.TrimTitle(titleValue);
        }

        internal string GetYear(IRow row)
        {
            return ExcelManager.GetValue(
              row,
              JsonManager.GetJsonTranslation(ExcelFileColumns.Year));
        }
    }
}
