using System;
using System.Collections.Generic;
using System.Linq;
using Logy.Entities.Engine;
using Logy.Entities.Localization;
using Logy.MwAgent.DotNetWikiBot;
using NPOI.HSSF.UserModel;

namespace Logy.Api.Mw.Excel
{
    public class EventsFromExcelManager
    {
        private const string TemplateName = "event";
        private const string TemplateExcelParameter = "excel";
        private const string TemplateExcelParsParameter = "excelParameters";
        private readonly string _title;
        private readonly string _fileNameAsPropertyName;

        public EventsFromExcelManager(ExcelImporter importer, HSSFRow row, Job job)
        {
            _title = importer.GetTitle(row);
            if (string.IsNullOrEmpty(_title))
            {
                return;
            }
            WikiPage = importer.NewPage(ref _title);

            WikiTemplateParametersModified = new List<string>();
            if (!WikiPage.Exists())
            {
                WikiPageCreated = true;
                WikiPage.Text = string.Format("{{{{{0}|{1}=}}}}", TemplateName, TemplateExcelParameter);
            }
            ThisExcelFile = "file:" + job.Template.Url;
            var excelValue = WikiPage.GetFirstTemplateParameter(TemplateName, TemplateExcelParameter);
            if (!excelValue.Contains(ThisExcelFile))
            {
                if (!string.IsNullOrEmpty(excelValue))
                {
                    const string ExcelDelimiter = ";";
                    excelValue += ExcelDelimiter;
                }
                excelValue += ThisExcelFile;
                SetTemplateValue(WikiPage, TemplateExcelParameter, excelValue, WikiTemplateParametersModified);
            }

            _fileNameAsPropertyName = job.Template.Url.Replace('.', ' ');

            foreach (var selectedColumn in (job.Parameters
                                            ?? string.Format(
                                                "{0};{1}",
                                                JsonManager.GetJsonTranslation(ExcelFileColumns.Description),
                                                JsonManager.GetJsonTranslation(ExcelFileColumns.Source))).Split(';'))
            {
                var value = importer.ExcelManager.GetValue(row, selectedColumn);
                string selectedColumnEn = null;
                foreach (Enum enumValue in typeof(ExcelFileColumns).GetEnumValues())
                {
                    if (JsonManager.GetJsonTranslation(enumValue)
                        .Equals(selectedColumn, StringComparison.InvariantCultureIgnoreCase))
                    {
                        selectedColumnEn = enumValue.ToString();
                        if (selectedColumnEn == ExcelFileColumns.Description.ToString())
                        {
                            value = ExcelManager.GetDescription(value);

                            var yearsInTitle = ExcelManager.GetYears(_title) ?? new[] { string.Empty };

                            int numY;
                            var date = importer.ExcelManager.GetValue(row, JsonManager.GetJsonTranslation(ExcelFileColumns.Year))
                                       ?? yearsInTitle[0];
                            var isCorrectDate = int.TryParse(date, out numY);
                            int numM = 1;
                            var month = importer.ExcelManager.GetValue(row, JsonManager.GetJsonTranslation(ExcelFileColumns.Month));
                            if (!string.IsNullOrEmpty(month))
                            {
                                date += "-" + month;
                                if (!int.TryParse(month, out numM))
                                    isCorrectDate = false;
                            }
                            int numD = 1;
                            var day = importer.ExcelManager.GetValue(row, JsonManager.GetJsonTranslation(ExcelFileColumns.Day));
                            if (!string.IsNullOrEmpty(day))
                            {
                                date += "-" + day;
                                if (!int.TryParse(day, out numD))
                                    isCorrectDate = false;
                            }
                            if (isCorrectDate)
                            {
                                var pageNameToAvoidDuplication = date + " " + WikiPage.Title;
                                var wasDate = WikiPage.GetTemplateParameter(TemplateName, "date");
                                if (wasDate.Count > 0 && wasDate[0] != date
                                    && ExcelManager.GetRowNumFromText(WikiPage.Text, _fileNameAsPropertyName) != row.RowNum.ToString())
                                {
                                    // this is duplication by Description
                                    var cell = importer.ExcelManager.GetCell(row, selectedColumn);
                                    cell.SetCellValue(pageNameToAvoidDuplication);
                                    importer.Import(row, job);
                                    return;
                                }
                                if (WikiPageCreated)
                                {
                                    // whether wikiPage is duplicate?
                                    var wikiPage2 = importer.NewPage(ref pageNameToAvoidDuplication);
                                    if (wikiPage2.Exists())
                                    {
                                        // do not save wikiPage, because there is wikiPage2
                                        WikiPage.Title = wikiPage2.Title;
                                    }
                                }

                                SetTemplateValue(
                                    WikiPage,
                                    "date",
                                    date,
                                    WikiTemplateParametersModified);
                                var shift = importer.ExcelManager.GetValue(row, JsonManager.GetJsonTranslation(ExcelFileColumns.Shift));
                                double shiftYears;
                                if (!string.IsNullOrEmpty(shift) && double.TryParse(shift, out shiftYears))
                                    SetTemplateValue(
                                        WikiPage,
                                        string.Format("{0}-redate", _fileNameAsPropertyName),
                                        new DateTime(numY, numM, numD).AddYears((int)shiftYears).ToString("yyyy-M-d"),
                                        WikiTemplateParametersModified);
                            }
                            else
                            {
                                SetTemplateValue(
                                    WikiPage,
                                    "dateText",
                                    yearsInTitle.Length > 1 ? yearsInTitle.Aggregate((c, s) => c + "-" + s) : date,
                                    WikiTemplateParametersModified);
                            }
                        }
                        else if (selectedColumnEn == ExcelFileColumns.Source.ToString())
                        {
                            var url = ExcelManager.GetUrl(value);
                            if (url != null)
                                SetTemplateValue(
                                    WikiPage,
                                    "url",
                                    url,
                                    WikiTemplateParametersModified);
                            var language = ExcelManager.GetLanguage(value);
                            if (language != null)
                                SetTemplateValue(
                                    WikiPage,
                                    "language",
                                    language,
                                    WikiTemplateParametersModified);
                            /// |nameAtUrl=<how to calc?>
                        }
                        else if (selectedColumnEn == ExcelFileColumns.Short.ToString())
                        {
                            var shortNamespace = job.Template.Url.Split(' ')[0];
                            var shortWikiPage = new Page(importer.WikiSite, string.Format("{0}:{1}", shortNamespace, value));
                            shortWikiPage.Load();
                            /* || !shortWikiPage.Text.Contains("interForm")) */
                            if (!shortWikiPage.Exists())
                            {
                                shortWikiPage.Save(shortWikiPage.Text + "{{interForm}}");
                            }
                            /*SetTemplateValue(
                                wikiPage,
                                string.Format("{0}-short", fileNameAsPropertyName),
                                "",
                                wikiTemplateParametersModified);
                            wikiPage.Text = wikiPage.Text.Replace(string.Format("|{0}-short=", fileNameAsPropertyName), null);*/

                            SetTemplateValue(
                                WikiPage,
                                "excelShorts",
                                shortNamespace + ":" + value,
                                WikiTemplateParametersModified,
                                null,
                                "short");
                            selectedColumnEn = null; // to disable Short parameter in event template
                        }
                        break;
                    }
                }
                if (selectedColumnEn != null)
                    SetTemplateValue(WikiPage, selectedColumnEn, value, WikiTemplateParametersModified, selectedColumn);
            }

            // do not set row above because 'date' parameter checks on duplication rows in GetRowNumFromText()
            SetTemplateValue(
                WikiPage,
                TemplateExcelParsParameter,
                row.RowNum.ToString(),
                WikiTemplateParametersModified,
                null,
                "row",
                "Count");
        }

        public string Title { get { return _title; } }
        public Page WikiPage { get; private set; }
        public string ThisExcelFile { get; private set; }
        internal List<string> WikiTemplateParametersModified { get; set; }
        internal bool WikiPageCreated { get; set; }
        
        private void SetTemplateValue(
            Page wikiPage,
            string parameterEn,
            string value,
            List<string> wikiTemplateParametersModified,
            string parameter = null,
            string separatorName = null,
            string countProperty = null)
        {
            if (separatorName != null)
                value = string.Format("[[{0}-{1}::{2}]]", _fileNameAsPropertyName, separatorName, value);
            if (!string.IsNullOrEmpty(value))
            {
                value = value.Trim();
                var wasValue = wikiPage.GetFirstTemplateParameter(TemplateName, parameterEn);
                ///&& (string.IsNullOrEmpty(wasValue)
                ///    || (wasValue != value && parameter != null/*overwrite only values from Excel, not calculated*/)

                var containedAsDelimited =
                    separatorName != null && !string.IsNullOrEmpty(wasValue) && wasValue.Contains(value);
                var separator = ';';
                if (separatorName != null && !string.IsNullOrEmpty(wasValue) && !wasValue.Contains(value))
                {
                    value = wasValue + separator + value;
                }

                if (wasValue != value && !containedAsDelimited)
                {
                    wikiPage.SetTemplateParameter(TemplateName, parameterEn, value, true);
                    wikiTemplateParametersModified.Add(parameter ?? parameterEn);
                    wasValue = value;
                }

                if (countProperty != null && separatorName != null)
                {
                    var valueParts = wasValue.Split(separator).Length;
                    /// if (valueParts > 1)
                    {
                        var countParameterName = parameterEn + countProperty;
                        var wasCountValue = wikiPage.GetFirstTemplateParameter(TemplateName, countParameterName);
                        var countValue = string.Format(
                            "[[{0}-{1}{2}::{3}]]",
                            _fileNameAsPropertyName,
                            separatorName,
                            countProperty,
                            valueParts);
                        if (wasCountValue != countValue)
                        {
                            wikiPage.SetTemplateParameter(TemplateName, countParameterName, countValue, true);
                            wikiTemplateParametersModified.Add(countParameterName);
                        }
                    }
                }
            }
        }
    }
}