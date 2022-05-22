using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataStatistics.Main;

/// <summary>
/// 工单模型
/// </summary>
public class OrderModel
{
    /// <summary>
    /// 工单辅助时间
    /// </summary>
    public DateOnly 首次处理时间 { get; set; }
    /// <summary>
    /// 服务橙
    /// </summary>
    public string 服务橙 { get; set; }
    /// <summary>
    /// 工单类型
    /// </summary>
    public string 工单类型 { get; set; }
    /// <summary>
    /// 工单来源
    /// </summary>
    public string 工单来源 { get; set; }
    public string 服务大区 { get; set; }

    public OrderModel()
    {
    }
}
