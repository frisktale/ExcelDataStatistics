using ExcelDataStatistics.Main;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Runtime.ConstrainedExecution;

namespace ExcelDataStatistics.Main;
public class Program
{
    public static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.Commercial;

        //获取配置
        var config = GetConfig();
        if (config == null)
        {
            Console.WriteLine("已生成配置文件，请填写配置文件后再运行程序");
            Console.ReadKey();
            return;
        }
        //获取数据所在的sheet
        var xlsxFile = new FileInfo(config.FileFullPath);
        if (!xlsxFile.Exists)
        {
            Console.WriteLine("无法找到导入数据表，请确认路径是否填写正确");
            Console.ReadKey();
            return;
        }
        using var package = new ExcelPackage(xlsxFile);
        var sheet = package.Workbook.Worksheets[config.DataSheetName];
        //读取数据
        List<OrderModel> list = ReadingData(sheet, config);


        var outputFIleinfo = new FileInfo(config.OutputPath);
        if (outputFIleinfo.Exists)
        {
            outputFIleinfo.Delete();
        }

        //创建sheet
        using var outputPackage = new ExcelPackage(config.OutputPath);

        //创建月汇总sheet
        GenerateMonthlySummarySheet(list, outputPackage);

        //创建每日汇总sheet
        GenerateDailySummarySheet(list, outputPackage);


        //保存excel
        outputPackage.Save();

        Console.WriteLine("统计成功");
        Console.ReadKey();

    }

    /// <summary>
    /// 设置Sheet的表头，包括表格标题，列名等
    /// </summary>
    /// <param name="outputCells">需要设置标题的sheet</param>
    /// <param name="mainTitleString">表头名称</param>
    /// <returns></returns>
    static int SetSheetTitle(ExcelWorksheet sheets, string mainTitleString = "工作记录表")
    {
        var cells = sheets.Cells;
        var mainTitle = cells["A1:F1"];
        mainTitle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        mainTitle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        mainTitle.Style.Font.Bold = true;
        mainTitle.Style.Font.Size = 18;
        mainTitle.Value = mainTitleString;
        mainTitle.Merge = true;
        cells["A2"].Value = "日期";
        cells["B2"].Value = "服务橙";
        cells["C2"].Value = "服务大区";
        cells["D2"].Value = "工单类型";
        cells["E2"].Value = "工单来源";
        cells["F2"].Value = "工单数";
        return 3;
    }

    /// <summary>
    /// 获取配置
    /// </summary>
    /// <returns></returns>
    static Config? GetConfig()
    {
        var sampleConfigFile = @"#工单记录表的路径，相对路径和绝对路径都支持
FileFullPath=E:\workspace\工单记录表.xlsx
#工单记录表中数据所在的sheet的名称
DataSheetName=数据导入
#数据的起始行（即除去表头的第一行）行号
StartRow=2
#数据的终止行行号
EndRow=280
#统计结果表的输出路径
OutputPath=E:\workspace\统计.xlsx
#首次跟进时间列名
HandingTimeColumnName=O
#工单来源列名
SourceColumnName=K
#工单类型列名
TypeColumnName=C
#工单区域
AreaColumn=E
#创建人列名
CreaterColumnName=L";
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddIniFile("config.ini", optional: false);
        if (!File.Exists("config.ini"))
        {
            File.WriteAllText("config.ini", sampleConfigFile);
            return null;
        }
        var root = builder.Build();
        var config = root.Get<Config>();
        return config;
    }

    /// <summary>
    /// 从sheet中读取数据
    /// </summary>
    /// <param name="sheet">数据源所在sheet</param>
    /// <param name="config">用户配置</param>
    /// <returns></returns>
    static List<OrderModel> ReadingData(ExcelWorksheet sheet, Config config)
    {
        var list = new List<OrderModel>();
        for (int i = config.StartRow; i <= config.EndRow; i++)
        {
            var cells = sheet.Cells;
            var orderModel = new OrderModel
            {
                工单类型 = cells[$"{config.TypeColumnName}{i}"].GetCellValue<string>(),
                工单来源 = cells[$"{config.SourceColumnName}{i}"].GetCellValue<string>(),
                服务橙 = cells[$"{config.CreaterColumnName}{i}"].GetCellValue<string>(),
                首次处理时间 = DateOnly.FromDateTime(cells[$"{config.HandingTimeColumnName}{i}"].GetCellValue<DateTime>()),
                服务大区 = cells[$"{config.AreaColumn}{i}"].GetCellValue<string>()
            };
            list.Add(orderModel);
        }

        return list;
    }

    /// <summary>
    /// 生成当月汇总信息
    /// </summary>
    /// <param name="list">从数据源读取的数据</param>
    /// <param name="package">导出的excel表</param>
    static void GenerateMonthlySummarySheet(List<OrderModel> list, ExcelPackage package)
    {
        var sheets = package.Workbook.Worksheets.Add("月汇总");
        var rowIndex = SetSheetTitle(sheets);
        var startIndex = rowIndex;
        //计算月汇总数据
        var groupByListWithoutDate = list.GroupBy(t => new { t.服务橙, t.工单类型, t.工单来源,t.服务大区 })
            .Select(t => new { t.Key.服务橙, t.Key.工单类型, t.Key.工单来源, Count = t.Count(),t.Key.服务大区 })
            .OrderBy(t => t.服务橙)
            .ThenBy(t => t.工单类型)
            .ThenBy(t => t.工单来源)
            .ThenBy(t => t.服务大区);

        var cells = sheets.Cells;
        //月汇总数据写入sheet
        foreach (var item in groupByListWithoutDate)
        {
            cells[rowIndex, 1].Value = "月汇总";
            cells[rowIndex, 2].Value = item.服务橙;
            cells[rowIndex, 3].Value = item.服务大区;
            cells[rowIndex, 4].Value = item.工单类型;
            cells[rowIndex, 5].Value = item.工单来源;
            cells[rowIndex, 6].Value = item.Count;
            rowIndex++;
        }

        //合并日期单元格
        var dateCellRange = cells[$"A{startIndex}:A{rowIndex - 1}"];
        dateCellRange.Merge = true;

        //所有单元格设为居中
        var style = cells[sheets.Dimension.Address].Style;
        style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        //设为自适应宽度
        cells[sheets.Dimension.Address].AutoFitColumns();
    }

    /// <summary>
    /// 生成每日汇总信息
    /// </summary>
    /// <param name="list">从数据源读取的数据</param>
    /// <param name="package">导出的excel表</param>
    static void GenerateDailySummarySheet(List<OrderModel> list, ExcelPackage package)
    {
        //计算每日汇总数据
        var groupByListWithDate = list.GroupBy(t => new { t.首次处理时间, t.服务橙, t.工单类型, t.工单来源, t.服务大区 })
            .Select(t => new { t.Key.首次处理时间, t.Key.服务橙, t.Key.工单类型, t.Key.工单来源, Count = t.Count(),t.Key.服务大区 })
            .OrderBy(t => t.首次处理时间)
            .ThenBy(t => t.服务橙)
            .ThenBy(t=>t.服务大区);

        //创建新sheet
        var sheets = package.Workbook.Worksheets.Add("单日汇总");


        //绘制表头
        var cells = sheets.Cells;
        var rowIndex = SetSheetTitle(sheets);
        var dateIndexList = new List<int>
        {
            rowIndex
        };
        DateOnly lastDate = groupByListWithDate.First().首次处理时间;
        //每日汇总数据写入sheet
        foreach (var item in groupByListWithDate)
        {
            cells[rowIndex, 1].Value = item.首次处理时间;
            cells[rowIndex, 2].Value = item.服务橙;
            cells[rowIndex, 3].Value = item.服务大区;
            cells[rowIndex, 4].Value = item.工单类型;
            cells[rowIndex, 5].Value = item.工单来源;
            cells[rowIndex, 6].Value = item.Count;
            if (lastDate.Day != item.首次处理时间.Day)
            {
                dateIndexList.Add(rowIndex);
                lastDate = item.首次处理时间;
            }
            rowIndex++;
        }
        dateIndexList.Add(rowIndex);

        //合并相同日期单元格
        for (int i = 0; i < dateIndexList.Count - 1; i++)
        {
            var dateCellRange2 = cells[$"A{dateIndexList[i]}:A{dateIndexList[i + 1] - 1}"];
            dateCellRange2.Merge = true;
        }

        //所有单元格设为居中
        var style = cells[sheets.Dimension.Address].Style;
        style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        //设为自适应宽度
        cells[sheets.Dimension.Address].AutoFitColumns();
    }
}
//设置许可
