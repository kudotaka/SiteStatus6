//using System.Collections.Generic;
//using System.Runtime.InteropServices;
using System.Collections.Immutable;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Graphics;

//using DocumentFormat.OpenXml.Drawing.Diagrams;
//using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<SiteStatusApp>();
app.Run();


public class SiteStatusApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<SiteStatusApp> logger;
    readonly IOptions<MyConfig> config;

    Dictionary<string, MySiteStatus> dicMySiteStatus = new Dictionary<string, MySiteStatus>();

    public SiteStatusApp(ILogger<SiteStatusApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Status(string definition, string progress, string save)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!File.Exists(definition))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{definition}");
            return;
        }
        if (!File.Exists(progress))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{progress}");
            return;
        }

        int definitionDataRow = config.Value.DefinitionDataRow;
        string definitionSheetName = config.Value.DefinitionSheetName;
        string definitionWordKeyToColum = config.Value.DefinitionWordKeyToColum;
        int progressDataRow = config.Value.ProgressDataRow;
        string progressSheetName = config.Value.ProgressSheetName;
        string progressWordKeyToColum = config.Value.ProgressWordKeyToColum;
        string progressIgnoreAtSiteKey = config.Value.ProgressIgnoreAtSiteKey;
        string vIPWordSiteKeyToStatus = config.Value.VIPWordSiteKeyToStatus;

        readDefinitionExcel(definition, definitionSheetName, definitionDataRow, definitionWordKeyToColum, dicMySiteStatus);
        convertZeroSiteName(dicMySiteStatus);
        readProgressExcel(progress, progressSheetName, progressDataRow, progressWordKeyToColum, progressIgnoreAtSiteKey, dicMySiteStatus);
        convertVIPStatus(vIPWordSiteKeyToStatus, dicMySiteStatus);

        printMySiteStatus(dicMySiteStatus);

        saveMySiteStatus(save, definition, progress, dicMySiteStatus);
//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての処理をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");
    }

    private void readDefinitionExcel(string excel, string sheetName, int firstDataRow, string wordKeyToColum, Dictionary<string, MySiteStatus> dic)
    {
        logger.ZLogInformation($"== start Definitionファイルの読み込み ==");
        bool isError = false;
        Dictionary<string, int> dicKeyToColumn = new Dictionary<string, int>();
        foreach (var keyAndValue in wordKeyToColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicKeyToColumn.Add(item[0], int.Parse(item[1]));
        }
        using FileStream fsExcel = new FileStream(excel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookExcel = new XLWorkbook(fsExcel);
        IXLWorksheets sheetsExcel = xlWorkbookExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsExcel)
        {
            if (sheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"excel:{excel},シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, {wordKeyToColum}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    MySiteStatus ss = new MySiteStatus();
                    foreach (var key in dicKeyToColumn.Keys)
                    {
                        var property = typeof(MySiteStatus).GetProperty(key);
                        if (property == null)
                        {
                            isError = true;
                            logger.ZLogError($"property is NULL  at sheet:{sheet.Name} row:{r} key:{key}");
                            continue;
                        }
                        IXLCell cellColumn = sheet.Cell(r, dicKeyToColumn[key]);
                        switch (cellColumn.DataType)
                        {
                            case XLDataType.DateTime:
                                property.SetValue(ss, cellColumn.GetValue<DateTime>().ToString("yyyy/MM/dd"));
                                break;
                            case XLDataType.Text:
                                if (key.Equals("workStatus"))
                                {
                                    property.SetValue(ss, convertWorkStatus(cellColumn.GetValue<string>()));
                                }
                                else
                                {
                                    property.SetValue(ss, cellColumn.GetValue<string>());
                                }
                                break;
                            case XLDataType.Number:
                                property.SetValue(ss, cellColumn.GetValue<int>().ToString());
                                break;
                            case XLDataType.Blank:
                                if (key.Equals("workStatus"))
                                {
                                    property.SetValue(ss, MyStatus.Work_InPrepare);
                                }
                                else
                                {
                                    logger.ZLogTrace($"cell is Blank type at sheet:{sheet.Name} row:{r}");
                                }
                                break;
                            default:
                                logger.ZLogError($"cell is NOT type ( DateTime | Text ) at sheet:{sheet.Name} row:{r}");
                                continue;
                        }
                    }
                    dic.Add(ss.siteKey, ss);
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] readDefinitionExcel()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] readDefinitionExcel()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end Definitionファイルの読み込み ==");
    }

    private void readProgressExcel(string excel, string sheetName, int firstDataRow, string wordKeyToColum, string progressIgnoreAtSiteKey, Dictionary<string, MySiteStatus> dic)
    {
        logger.ZLogInformation($"== start Progressファイルの読み込み ==");
        bool isError = false;
        List<string> listProgressIgnoreAtSiteKey = new List<string>();
        foreach (var ignore in progressIgnoreAtSiteKey.Split(','))
        {
            listProgressIgnoreAtSiteKey.Add(ignore);
        }
        Dictionary<string, int> dicKeyToColumn = new Dictionary<string, int>();
        foreach (var keyAndValue in wordKeyToColum.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicKeyToColumn.Add(item[0], int.Parse(item[1]));
        }
        using FileStream fsExcel = new FileStream(excel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookExcel = new XLWorkbook(fsExcel);
        IXLWorksheets sheetsExcel = xlWorkbookExcel.Worksheets;
        foreach (IXLWorksheet? sheet in sheetsExcel)
        {
            if (sheetName.Equals(sheet.Name))
            {
                int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                logger.ZLogInformation($"excel:{excel},シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, {wordKeyToColum}");

                for (int r = firstDataRow; r < lastUsedRowNumber + 1; r++)
                {
                    string siteKey = "";
                    string statusWord = "";
                    foreach (var key in dicKeyToColumn.Keys)
                    {
                        IXLCell cellColumn = sheet.Cell(r, dicKeyToColumn[key]);
                        switch (cellColumn.DataType)
                        {
                            case XLDataType.Text:
                                if (key.Equals("siteKey"))
                                {
                                    siteKey = cellColumn.GetValue<string>();
                                }
                                else if (key.Equals("status"))
                                {
                                    statusWord = cellColumn.GetValue<string>();
                                }
                                break;
                            case XLDataType.Blank:
                                isError = true;
                                logger.ZLogError($"cell is Blank type at siteKey:{siteKey} sheet:{sheet.Name} row:{r}");
                                break;
                            default:
                                isError = true;
                                logger.ZLogError($"cell is NOT type ( Text ) at siteKey:{siteKey} sheet:{sheet.Name} row:{r}");
                                continue;
                        }
                    }
                    if (listProgressIgnoreAtSiteKey.Contains(siteKey))
                    {
                        logger.ZLogTrace($"[readProgressExcel] 除外しました({siteKey})");
                        continue;
                    }
                    if (!dicMySiteStatus.ContainsKey(siteKey))
                    {
                        isError = true;
                        logger.ZLogError($"[ERROR] progressの拠点番号({siteKey})がdefinitionの拠点番号と一致するものがありませんでした");
                        continue;
                    }
                    else
                    {
                        dic[siteKey].status = convertProgressStatus(statusWord);
                    }
                }
            }
            else
            {
                logger.ZLogTrace($"Miss {sheet.Name}");
            }
        }

        if (!isError)
        {
            logger.ZLogInformation($"[OK] readProgressExcel()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] readProgressExcel()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end Progressファイルの読み込み ==");
    }

    private void convertVIPStatus(string vIPWordSiteKeyToStatus, Dictionary<string, MySiteStatus> dic)
    {
        Dictionary<string, int> dicSiateKeyToStatus = new Dictionary<string, int>();
        foreach (var keyAndValue in vIPWordSiteKeyToStatus.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicSiateKeyToStatus.Add(item[0], int.Parse(item[1]));
        }

        foreach (var key in dicSiateKeyToStatus.Keys)
        {
            if (dic.ContainsKey(key))
            {
                dic[key].status = (MyStatus)dicSiateKeyToStatus[key];
            }
        }
    }

    private void saveMySiteStatus(string save, string definition, string progress, Dictionary<string, MySiteStatus> dic)
    {
        logger.ZLogInformation($"== start ファイルの新規作成 ==");
        bool isError = false;

        const int SAVE_COLUMN_INPUTDATA = 1;
        const int SAVE_ROW_INPUTDATA = 1;
        const int SAVE_COLUMN_SITEKEY = 1;
        const int SAVE_COLUMN_SITENAME = 2;
        const int SAVE_COLUMN_WORK_STATUS = 3;
        const int SAVE_COLUMN_STATUS = 4;
        const int SAVE_FIRST_ROW = SAVE_ROW_INPUTDATA + 4;
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("status");

        worksheet.Cell(SAVE_ROW_INPUTDATA, SAVE_COLUMN_INPUTDATA).SetValue(convertDateTimeToJst(DateTime.Now)).Style.Font.SetFontName("Meiryo UI");
        worksheet.Cell(SAVE_ROW_INPUTDATA + 1, SAVE_COLUMN_INPUTDATA).SetValue(definition).Style.Font.SetFontName("Meiryo UI");
        worksheet.Cell(SAVE_ROW_INPUTDATA + 2, SAVE_COLUMN_INPUTDATA).SetValue(progress).Style.Font.SetFontName("Meiryo UI");

        worksheet.Cell(SAVE_FIRST_ROW, SAVE_COLUMN_SITEKEY).SetValue("拠点キー").Style.Font.SetFontName("Meiryo UI");
        worksheet.Cell(SAVE_FIRST_ROW, SAVE_COLUMN_SITENAME).SetValue("拠点名").Style.Font.SetFontName("Meiryo UI");
        worksheet.Cell(SAVE_FIRST_ROW, SAVE_COLUMN_WORK_STATUS).SetValue("工事ステータス").Style.Font.SetFontName("Meiryo UI");
        worksheet.Cell(SAVE_FIRST_ROW, SAVE_COLUMN_STATUS).SetValue("ステータス").Style.Font.SetFontName("Meiryo UI");

        int row = SAVE_FIRST_ROW + 1;
        var keys = dic.Keys.ToImmutableList().Sort();
        foreach (var key in keys)
        {
            MySiteStatus ss = dic[key];
            worksheet.Cell(row, SAVE_COLUMN_SITEKEY).SetValue(ss.siteKey).Style.Font.SetFontName("Meiryo UI");
            worksheet.Cell(row, SAVE_COLUMN_SITENAME).SetValue(ss.siteName).Style.Font.SetFontName("Meiryo UI");
            worksheet.Cell(row, SAVE_COLUMN_WORK_STATUS).SetValue(convertStatusToReadableStatus(ss.workStatus)).Style.Font.SetFontName("Meiryo UI");
            worksheet.Cell(row, SAVE_COLUMN_STATUS).SetValue(convertStatusToReadableStatus(ss.status)).Style.Font.SetFontName("Meiryo UI");
            row++;
        }

        worksheet.Column(SAVE_COLUMN_SITEKEY).AdjustToContents();
        worksheet.Column(SAVE_COLUMN_SITENAME).AdjustToContents();
        worksheet.Column(SAVE_COLUMN_WORK_STATUS).AdjustToContents();
        worksheet.Column(SAVE_COLUMN_STATUS).AdjustToContents();
        workbook.SaveAs(save);
        if (!isError)
        {
            logger.ZLogInformation($"[OK] saveMySiteStatus()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] saveMySiteStatus()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end ファイルの新規作成 ==");
    }

    private void convertZeroSiteName(Dictionary<string, MySiteStatus> dic)
    {
        foreach (var value in dic.Values)
        {
            value.siteName = convertSiteNameToZeroSiteName(value.siteName);
        }
    }

    private MyStatus convertWorkStatus(string status)
    {
        string wordStringToWorkStatus = config.Value.WordStringToWorkStatus;
        Dictionary<string, int> dicStringToWorkStatus = new Dictionary<string, int>();
        foreach (var keyAndValue in wordStringToWorkStatus.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicStringToWorkStatus.Add(item[0], int.Parse(item[1]));
        }

        if (dicStringToWorkStatus.ContainsKey(status))
        {
            return (MyStatus)dicStringToWorkStatus[status];
        }
        return MyStatus.Work_InPrepare;
    }

    private MyStatus convertProgressStatus(string status)
    {
        string wordStringToStatus = config.Value.WordStringToStatus;
        Dictionary<string, int> dicStringToStatus = new Dictionary<string, int>();
        foreach (var keyAndValue in wordStringToStatus.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicStringToStatus.Add(item[0], int.Parse(item[1]));
        }

        if (dicStringToStatus.ContainsKey(status))
        {
            return (MyStatus)dicStringToStatus[status];
        }
        return MyStatus.UnKnown;
    }

    private void printMySiteStatus(Dictionary<string, MySiteStatus> dic)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var site in dic.Values.ToList())
        {
            logger.ZLogTrace($"キー:{site.siteKey},拠点名:{site.siteName},工事ステータス:{convertStatusToReadableStatus(site.workStatus)},ステータス:{convertStatusToReadableStatus(site.status)}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }

    private string convertStatusToReadableStatus(MyStatus status)
    {
        switch (status)
        {
            case MyStatus.Init:
                return "着手前";
            case MyStatus.Survey_Scheduling:
                return "調査|日程調整中";
            case MyStatus.Survey_Scheduled:
                return "調査|日程確定";
            case MyStatus.Survey_Completed:
                return "調査|完了";
            case MyStatus.Work_Scheduling:
                return "工事|日程調整中";
            case MyStatus.Work_Scheduled:
                return "工事|日程確定";
            case MyStatus.Work_InPrepare:
                return "工事|準備";
            case MyStatus.Work_InProgress:
                return "工事|工事中";
            case MyStatus.Work_Tested:
                return "工事|NW試験完了";
            case MyStatus.Work_Completed:
                return "工事|全完了";
            case MyStatus.Create_CompleteBook:
                return "図書作成";
            case MyStatus.Finish:
                return "すべて完了";
            case MyStatus.VIP:
                return "特別対応";
            case MyStatus.NotAvailable:
                return "廃止・対象外 等";
            case MyStatus.UnKnown:
                return "不明";
            default:
                break;
        }
        return status.ToString();
    }

    private string convertSiteNameToZeroSiteName(string target)
    {
        int index = target.IndexOf('-');
        switch (index)
        {
            case 1:
                return "000"+target;
            case 2:
                return "00"+target;
            case 3:
                return "0"+target;
            default:
                break;
        }
        int index2 = target.IndexOf('_');
        switch (index2)
        {
            case 1:
                return "000"+target;
            case 2:
                return "00"+target;
            case 3:
                return "0"+target;
            default:
                break;
        }
        return target;
    }

    private string convertDateTimeToDateAndDayofweek(DateTime day)
    {
        StringBuilder sb = new StringBuilder();
        switch (day.DayOfWeek)
        {
        case DayOfWeek.Sunday:
            sb.Append(day.ToString("yyyy/MM/dd(日)"));
            break;
        case DayOfWeek.Monday:
            sb.Append(day.ToString("yyyy/MM/dd(月)"));
            break;
        case DayOfWeek.Tuesday:
            sb.Append(day.ToString("yyyy/MM/dd(火)"));
            break;
        case DayOfWeek.Wednesday:
            sb.Append(day.ToString("yyyy/MM/dd(水)"));
            break;
        case DayOfWeek.Thursday:
            sb.Append(day.ToString("yyyy/MM/dd(木)"));
            break;
        case DayOfWeek.Friday:
            sb.Append(day.ToString("yyyy/MM/dd(金)"));
            break;
        case DayOfWeek.Saturday:
            sb.Append(day.ToString("yyyy/MM/dd(土)"));
            break;
        }
        return sb.ToString();
    }

    private string convertDateTimeToJst(DateTime day)
    {
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, jstTimeZoneInfo).ToString("yyyy/MM/dd HH:mm");
    }


/*
    private string replaceDateTimeString(string dateTimeString)
    {
        return dateTimeString.Replace(" ","").Replace("、",",").Replace(",","|");
    }
*/
}

//==
public class MyConfig
{
    public int DefinitionDataRow {get; set;} = -1;
    public string DefinitionSheetName {get; set;} = "";
    public string DefinitionWordKeyToColum {get; set;} = "";
    public string WordStringToWorkStatus {get; set;} = "";
    public int ProgressDataRow {get; set;} = -1;
    public string ProgressSheetName {get; set;} = "";
    public string ProgressWordKeyToColum {get; set;} = "";
    public string ProgressIgnoreAtSiteKey {get; set;} = "";
    public string WordStringToStatus {get; set;} = "";
    public string VIPWordSiteKeyToStatus {get; set;} = "";
}

public enum MyStatus
{
    Init = 0,
    Survey_Scheduling = 10,
    Survey_Scheduled = 11,
    Survey_Completed = 12,
    Work_Scheduling = 20,
    Work_Scheduled = 21,
    Work_InPrepare = 24,
    Work_InProgress = 25,
    Work_Tested = 26,
    Work_Completed = 29,
    Create_CompleteBook = 40,
    Finish = 50,
    VIP = 80,
    NotAvailable = 90,
    UnKnown = 91
}

public class MySiteStatus
{
    public string siteKey { set; get; } = "";
    public string siteNumber { set; get; } = "";
    public string siteName { set; get; } = "";
    public MyStatus workStatus { set; get; } = MyStatus.UnKnown;
    public MyStatus status = MyStatus.UnKnown;
}