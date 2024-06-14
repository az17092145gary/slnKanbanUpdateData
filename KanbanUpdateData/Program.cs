
using System.Net.Mail;
using System.Net;
using KanbanUpdateData.Model;
using System.Data.SqlClient;
using Dapper;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;
using KanbanUpdateData.Modol;
using System.ServiceProcess;
using System.ComponentModel.Design;
using System.ServiceProcess;



string connectionString = "Data Source= 192.168.0.82;Initial Catalog=AIOT;Persist Security Info=True;User ID=sa;Password=P@ssw0rd;Encrypt=True;TrustServerCertificate=True";
string produndtConnectionString = "User ID=ds;Password=ds;Data Source=192.168.160.207:1521/topprod";

//執行
try
{
    executeMethod();
}
catch
{
    DateTime dateTime = DateTime.Now;
    string content = "KanBan程式執行有誤，時間 : " + dateTime;
    SendEMail(content, false);
}

async Task<TempData> createTemp(LOWDATA item)
{
    TempData tempData = new TempData();
    tempData.StopTenUP = new List<MachineStop>();
    tempData.StopTenDown = new List<MachineStop>();
    var Date = item.Time.Split(' ');
    tempData.Line = item.ProductLine;
    //tempData.Factory = item.Factory;
    tempData.State = "未自動運行";
    tempData.ModelStartTime = item.Time;
    tempData.Model = item.Model;
    tempData.DeviceOrder = item.DeviceOrder;
    tempData.Date = Date[0];
    tempData.DeviceName = item.DeviceName;
    tempData.Item = item.Item;
    tempData.Sum = 0;
    tempData.NGSum = 0;
    tempData.Product = item.Product;
    tempData.Activation = item.Activation;
    tempData.Throughput = item.Throughput;
    tempData.Defective = item.Defective;
    tempData.Exception = item.Exception;
    tempData.QIMSuperMode = false;
    tempData.DMISuperMode = false;
    tempData.MTCSuperMode = false;
    return tempData;
}
string findWKC(SqlConnection con, string sql, string DeviceName, string strTime, string endTime, DateTime lastTimeData)
{
    strTime = Convert.ToDateTime(strTime).AddDays(-1).ToString("yyyy-MM-dd 08:00:00");
    endTime = Convert.ToDateTime(strTime).AddDays(1).ToString("yyyy-MM-dd 08:00:00");
    sql = $"SELECT Min(TIME) FROM [AIOT].[dbo].[Machine_Data]";
    if (Convert.ToDateTime(strTime) < lastTimeData)
    {
        return "找不到該工單";
    }
    else
    {
        sql = $"SELECT TOP(1) VALUE FROM [AIOT].[dbo].[Machine_Data] where NAME like '%WKC%' and time between '{strTime}' and '{endTime}' AND Quality <> 'Bad' AND DEVICENAME = '{DeviceName}' AND VALUE <> '' AND VALUE IS NOT NULL  order by TIME desc";
        var data = con.QueryFirstOrDefault<string>(sql);
        if (data != null && data.TrimStart().TrimEnd().Length == 13)
        {
            return data.TrimStart().TrimEnd();

        }
        else
        {
            data = findWKC(con, sql, DeviceName, strTime, endTime, lastTimeData);
            return data.TrimStart().TrimEnd();
        }
    }
}
void inputData(out string sql, string _strTime, string _endTime, string _Date, SqlConnection conn, List<TempData> completeLowDatas, List<NonWork> completeNonWorkDataS, List<DailyERRData> dailyERRDatas)
{

    foreach (var item in completeLowDatas)
    {
        if (string.IsNullOrEmpty(item.EndTime))
        {
            item.EndTime = _endTime;
            item.SumTime = (Convert.ToDateTime(item.EndTime) - Convert.ToDateTime(item.StartTime)).TotalMinutes;

        }
    }
    //判斷是否有沒有Endtime的NonWorkData
    foreach (var item in completeNonWorkDataS)
    {
        if (string.IsNullOrEmpty(item.EndTime))
        {
            item.EndTime = _endTime;
            item.SumTime = (Convert.ToDateTime(item.EndTime) - Convert.ToDateTime(item.StartTime)).TotalMinutes;
        }
    }



    //排除150R試做工單
    var machineDataList = completeLowDatas.Where(x => !x.WorkCode.Contains("150R-")).GroupBy(x =>
    new { x.WorkCode, x.DeviceOrder, x.Item, x.Product, x.Line, x.DeviceName, x.Activation, x.Throughput, x.Defective, x.Exception }).Select(y =>
    {
        var EndTime = y.Max(x => x.EndTime);
        var tempModelSumTime = y.Max(x => x.ModelStartTime);
        var ModelSumTime = tempModelSumTime != null ? Math.Round((Convert.ToDateTime(EndTime) - Convert.ToDateTime(tempModelSumTime)).TotalMinutes, 2) : 0;
        return new
        {
            y.Key.WorkCode,
            y.Key.Line,
            y.Key.Item,
            y.Key.Product,
            //y.Key.Factory,
            y.Key.DeviceOrder,
            y.Key.DeviceName,
            y.Key.Activation,
            y.Key.Throughput,
            y.Key.Defective,
            y.Key.Exception,
            State = y.OrderByDescending(x => x.EndTime).FirstOrDefault()?.State,
            ModelSumTime,
            StartTime = y.Min(x => x.StartTime),
            SumTime = y.Sum(z => z.SumTime),
            Sum = y.Sum(z => Convert.ToDouble(z.Sum)),
            //如果關機會重複計算NGS，所以取最大的NGS
            NGS = y.Max(z => Convert.ToDouble(z.NGSum)),
            StopRunTime = y.Sum(z => Convert.ToDouble(z.RunSumStopTime))
        };
    }).ToList();
    //資料庫取出標準產能
    #region

    sql = $"SELECT * FROM [AIOT].[dbo].[Standard_Production_Efficiency_Benchmark]";
    var stdPerformanceList = conn.Query<StdPerformance>(sql).ToList();
    //標準產能(依照當日執行時間最長的工單抓取標準產能)
    var pcsList = new List<TempStd>();
    //全部品名
    var Product_NameList = new List<TempStd>();

    var tempStdList = completeLowDatas.Where(x => !x.WorkCode.Contains("150R-")).GroupBy(x => new { x.Item, x.Product, x.Line, x.WorkCode }).Select(y => new
    {
        //y.Key.Factory,
        y.Key.WorkCode,
        y.Key.Product,
        y.Key.Line,
        y.Key.Item,
        Model = y.Select(x => x.Model).FirstOrDefault(),
    }).ToList();
    //工單 搜尋 ERP資料庫找到對應的料件，料件對應LOCAL DB 找出品名、PCS
    string emailContext = "";
    using (OracleConnection oracleConnection = new OracleConnection(produndtConnectionString))
    {

        foreach (var item1 in tempStdList)
        {
            var partsql = $"select sfb05 as Part_No from dipf2.sfb_file,dipf2.ima_file where sfb05 =ima01 and sfb87='Y' and sfb01 = '{item1.WorkCode}'";
            var part = oracleConnection.Query<string>(partsql).FirstOrDefault();

            //2024/04/03增加移除找不到的PCS並寄信
            if (part != null)
            {
                var PCS = stdPerformanceList.Where(x => x.Part_No == part.ToString() && x.Model == item1.Model).Select(x => x.PCS).FirstOrDefault();
                if (PCS == null)
                {
                    var tempList = machineDataList.Where(x => x.Line == item1.Line && x.Item == item1.Item && x.Product == item1.Product).Select(x => x).ToList();
                    machineDataList = machineDataList.Except(tempList).ToList();

                    //異常時停止KanBan自動運行Service
                    ServiceController sc = new ServiceController("AIoTKanBanService");
                    if (!(sc.Status == ServiceControllerStatus.Stopped) || (sc.Status == ServiceControllerStatus.StopPending))
                    {
                        sc.Stop();
                        sc.Refresh();
                    }
                    emailContext += "日期:" + _Date + "\n" + "工單編號:" + $"{item1.WorkCode}該工單對應不到標準產能。";
                    SendEMail(emailContext, true);
                    break;

                }
                else
                {
                    var data = new TempStd();
                    data.WorkCode = item1.WorkCode;
                    data.PCS = PCS;
                    data.Item = item1.Item;
                    data.Line = item1.Line;
                    data.Product = item1.Product;
                    //data.Factory = item1.Factory;
                    data.Product_Name = stdPerformanceList.Where(x => x.Part_No == part.ToString()).Select(x => x.Product_Name).FirstOrDefault();
                    pcsList.Add(data);
                }
            }
            else
            {
                var tempList = machineDataList.Where(x => x.Line == item1.Line && x.Item == item1.Item && x.Product == item1.Product).Select(x => x).ToList();
                machineDataList = machineDataList.Except(tempList).ToList();
                //異常時停止KanBan自動運行Service
                ServiceController sc = new ServiceController("AIoTKanBanService");
                if (!(sc.Status == ServiceControllerStatus.Stopped) || (sc.Status == ServiceControllerStatus.StopPending))
                {
                    sc.Stop();
                    sc.Refresh();
                }
                emailContext += "日期:" + _Date + "\n" + "工單編號:" + $"{item1.WorkCode}請確認該工單編號的正確性，於ERP無對應工單編號。";
                SendEMail(emailContext, true);
            }
        }


    }
    #endregion
    #region
    //排除PAT
    dailyERRDatas = dailyERRDatas.Where(x => x.Type != "PAT").Select(y => y).ToList();

    #endregion
    //統整機台稼動率、良率、生產產能
    #region
    //排除切單機，依照線號進行分組
    //MCT:最大產能工時:24小時
    //SC:標準產能(組裝機)
    //ETC:預計投入工時(組裝機)
    //PT:計畫開機工時
    //ACT:實際投入工時
    //ACTH:實際產出工時
    //YieIdAO:測試機產出量
    //AO:實際產出數量(全檢機)
    //CAPU:產能利用率
    //ADR:時間稼動率

    //依照 廠區、產品、產線統整
    var Line_MachineDailyData = machineDataList.GroupBy(x => new { x.WorkCode, x.Line, x.Product, x.Item }).Select(y =>
    {


        var lastDeviceOrder = Convert.ToInt32(completeLowDatas.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Item == y.Key.Item && x.Product == y.Key.Product).Max(x => x.DeviceOrder));

        //工單開始時間
        var StartTime = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.StartTime).FirstOrDefault();
        var ModelSumTime = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.ModelSumTime).FirstOrDefault().ToString();

        //關機時間
        //var NonCloseTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder && x.Line == y.Key.Line && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("closeMachine")).Sum(x => x.SumTime) / 60;
        //預計投入工時(Throughput)
        var ETC = Math.Round((double)y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.SumTime).FirstOrDefault(0.0) / 60, 2);
        // 6S(Throughput)
        //var Non6sTime = dailyNonWorkDatas.Where(x =>x.WorkCode == y.Key.WorkCode && x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("6S")).Select(x => x.SumCount).FirstOrDefault() / 60;
        //缺料停機(Throughput)
        var NonDMITime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder && x.Line == y.Key.Line && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("DMI")).Sum(x => x.SumTime) / 60;
        // 品檢模式(Throughput)
        var NonQIMTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder && x.Line == y.Key.Line && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("QIM")).Sum(x => x.SumTime) / 60;

        //退出品檢模式到缺料停機或機台改善之前時間(Throughput)
        var NonStopQTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder && x.Line == y.Key.Line && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("StopQTime")).Sum(x => x.SumTime) / 60;

        //無開機工時
        var NonTime = Math.Round(NonDMITime + NonQIMTime + NonStopQTime, 2);
        //設備損失工時已排除(Exception)//自動運行狀態下臨停10分鐘以上
        var StopRunTime = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.StopRunTime).FirstOrDefault(0.0) / 60;
        //機台故障維修//人員操作機故障時間
        var MTCTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder && x.Line == y.Key.Line && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("MTC")).Sum(x => x.SumTime) / 60;
        StopRunTime = Math.Round(StopRunTime + MTCTime, 2);

        var PT = Math.Round((double)(ETC - NonTime), 2);
        var ACT = Math.Round(PT - StopRunTime, 2);
        //顯示當日更換的所有品名
        //var Product_Name = string.Join(", ", Product_NameList.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).GroupBy(x => x.Product_Name).Select(x => x.Key));
        var Product_Name = pcsList.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Item == y.Key.Item && x.Product == y.Key.Product).Select(y => y.Product_Name).FirstOrDefault();
        var SC = Convert.ToDouble(pcsList.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Item == y.Key.Item && x.Product == y.Key.Product).Select(y => y.PCS).FirstOrDefault());
        //實際產量(Throughput)
        var AO = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.Sum).FirstOrDefault(0.0);
        //var lastYieIDAO = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.NGS).FirstOrDefault(0.0);
        //AO = AO + lastYieIDAO;
        var ACTH = Math.Round((AO / SC), 2);
        var YieIdAO = y.Where(x => x.Defective == true && x.Throughput != true).Select(x => x.Sum).FirstOrDefault(0.0);
        //var AllNGS = y.Where(x => x.Defective == true).Select(x => x.NGS).DefaultIfEmpty(0.0).Sum();
        var AllNGS = Convert.ToDouble(dailyERRDatas.Where(x => x.Type == "NGI" && x.Line == y.Key.Line && x.WorkCode == y.Key.WorkCode).Sum(x => x.Count));
        //var tempPerformance = Math.Round(((AO / SC) / ACT) * 100, 2);
        var tempPerformance = Math.Round((ACTH / ACT) * 100, 2);
        var Performance = (tempPerformance > 100 ? 99 : tempPerformance).ToString();
        //(測試機 + 全檢(不良數)) / 測試產出量 * 100
        var YieId = Math.Round((100 - (AllNGS / YieIdAO) * 100), 2);
        YieId = double.IsNaN(YieId) || double.IsNegativeInfinity(YieId) || double.IsPositiveInfinity(YieId) || double.IsNegative(YieId) ? 100 : YieId;
        var Availability = Math.Round(((ACT / PT) * 100), 2).ToString();
        var OEE = Math.Round((Convert.ToDouble(Performance) / 100) * (Convert.ToDouble(YieId) / 100) * (Convert.ToDouble(Availability) / 100) * 100, 2).ToString();
        //工單總臨停
        var StopCount = dailyERRDatas.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Type == "ERR").Sum(x => x.Count);
        var allTime = ACT >= 1.00 ? ACT : 1.0;
        //平均臨停
        var AVGStopCount = Math.Round(Convert.ToDouble(StopCount / allTime / lastDeviceOrder), 2).ToString();
        //最後一台機目前運行狀態
        var State = y.Where(x => Convert.ToInt32(x.DeviceOrder) == lastDeviceOrder).Select(x => x.State).FirstOrDefault("未自動運行");
        return new
        {
            ModelSumTime,
            NonDMITime,
            NonQIMTime,
            NonStopQTime,
            lastDeviceOrder,
            Product_Name,
            y.Key.WorkCode,
            //y.Key.Factory,
            y.Key.Item,
            y.Key.Product,
            State,
            y.Key.Line,
            Date = _Date,
            StartTime,
            MCT = "24",
            SC,
            ETC,
            PT,
            ACT,
            ACTH,
            AO,
            CAPU = Math.Round(((double)(ETC / 24) * 100), 2),
            ADR = Math.Round((double)((PT / ETC) * 100), 2),
            Performance,
            YieId,
            Availability,
            OEE,
            NonTime,
            StopRunTime,
            AVGStopCount,
            AllNGS
        };
    });

    //2024/04/19 排除異常數據

    Line_MachineDailyData = Line_MachineDailyData.Where(x => x.AO > 0 && Convert.ToDouble(x.Performance) > 0 && x.SC > 0).Select(x => x).ToList();
    //刪除看板系統紀錄
    sql = " DELETE [AIOT].[dbo].[KanBan_Line_MachineData] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_Machine_StopTenDown] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_Machine_StopTenUp] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_MachineERRData] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_MachineNonTime] ";
    conn.Execute(sql);
    sql = @" insert into [AIOT].[dbo].[KanBan_Line_MachineData]
                            Values(@Item,@Product,@State,@WorkCode,@ModelSumTime,@Product_Name,@Line,@Date,@StartTime,@MCT,@SC,@ETC,@PT,@ACT,@ACTH,@AO,@CAPU,@ADR,@Performance,@YieId,@Availability,@OEE,@NonTime,@StopRunTime,@AVGStopCount,@AllNGS)";
    ////執行資料庫
    conn.Execute(sql, Line_MachineDailyData);
    //無開機資料匯入資料庫
    sql = @"insert into [AIOT].[dbo].[KanBan_MachineNonTime]
                            Values(@DeviceName,@WorkCode,@Description,@Name,@Date,@StartTime,@EndTime,@SumTime)";
    ////執行資料庫
    conn.Execute(sql, completeNonWorkDataS);
    //#endregion
    ////統整錯誤訊息、10分鐘以上與10分鐘以下開機時間紀錄
    //#region
    //
    var TenUpDatas = new List<MachineStop>();
    var TenDownDatas = new List<MachineStop>();
    foreach (var item in completeLowDatas)
    {
        TenUpDatas.AddRange(item.StopTenUP);
        TenDownDatas.AddRange(item.StopTenDown);
    }
    //錯誤訊息DailyERRDatas匯入資料庫
    sql = @"insert into [AIOT].[dbo].[KanBan_MachineERRData]
                            Values(@DeviceName,@WorkCode,@Date,@Time,@Type,@Name,@Count)";
    //執行資料庫
    conn.Execute(sql, dailyERRDatas);

    //10分鐘以上開機時間紀錄匯入資料庫
    sql = "INSERT INTO [AIOT].[dbo].[KanBan_Machine_StopTenUp] VALUES(@DeviceName,@WorkCode,@Date,@StartTime,@EndTime,@SumTime)";
    //執行資料庫
    conn.Execute(sql, TenUpDatas);

    //10分鐘以下開機時間紀錄匯入資料庫
    sql = "INSERT INTO [AIOT].[dbo].[KanBan_Machine_StopTenDown] VALUES(@DeviceName,@WorkCode,@Date,@StartTime,@EndTime,@SumTime)";
    //執行資料庫
    conn.Execute(sql, TenDownDatas);
    #endregion
    //string context = _Date + "資料已產出，如資料有異常請通知資訊";
    //SendEMail(context);
}

async void executeMethod()
{
    //開始時間
    var _strTime = DateTime.Today.Add(new TimeSpan(8, 0, 0)).ToString("yyyy-MM-dd HH:mm:ss");
    //結束時間
    var _endTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    //今日日期
    //_strTime = "2024-05-31 08:00:00";
    //_endTime = "2024-05-31 13:47:00";
    var _Date = Convert.ToDateTime(_strTime).ToString("yyyy-MM-dd");
    //取出當天的LowData
    //string sql = @"SELECT F.Factory, F.Item,F.Product,F.Model,F.DeviceOrder,F.ProductLine,F.Activation,F.Throughput,F.Defective,F.Exception,MD.DeviceName,MD.NAME,MD.QUALITY,MD.TIME,MD.VALUE,MD.Description ";
    //sql += " FROM(select * FROM[AIOT].[dbo].[Machine_Data] WHERE TIME BETWEEN @_strTime AND @_endTime ) AS MD";
    //sql += " LEFT JOIN [AIOT].[dbo].[Factory]as F ON F.[IODviceName] = MD.[DeviceName] ";
    ////sql += " Where F.ProductLine = '12' OR F.ProductLine = '10'";
    //sql += " ORDER BY TIME,NAME";


    string sql = @"SELECT MD.DeviceName
    ,MD.NAME
    ,MD.QUALITY
    ,MD.TIME
    ,MD.VALUE
    ,MD.Description
    ,I.ItemName as Item
    ,PD.ProductName as Product
    ,L.LineName as ProductLine
    ,M.Model
    ,M.Defective
    ,M.Activation
    ,M.Exception
    ,M.Throughput
    ,M.DeviceOrder
      FROM (SELECT * FROM[AIOT].[dbo].[Machine_Data] WHERE TIME BETWEEN @_strTime AND @_endTime)  AS MD
      LEFT JOIN [AIOT].[dbo].[Machine] AS M ON MD.DeviceName = M.IODviceName
      LEFT JOIN [AIOT].[dbo].[ProductProductionLines] AS PP ON PP.id = M.ProductProductionLinesID
      LEFT JOIN [AIOT].[dbo].[ProductLine] AS L ON PP.LineID = L.LineID
      LEFT JOIN [AIOT].[dbo].[Product] AS PD ON PD.ProductID = PP.ProductID
      LEFT JOIN [AIOT].[dbo].[Item] AS I ON I.ItemID = PD.ItemID
      ORDER BY TIME,NAME";


    //暫存資料分類
    var tempLowDatas = new List<TempData>();
    //完成資料分類
    var completeLowDatas = new List<TempData>();
    //暫存無開機工時分類
    var tempNonWorkData = new List<NonWork>();
    //完成無開機工時分類
    var completeNonWorkDataS = new List<NonWork>();
    //異常碼
    var dailyERRDatas = new List<DailyERRData>();



    using (SqlConnection conn = new SqlConnection(connectionString))
    {
        Stopwatch stopwatch1 = new Stopwatch();
        stopwatch1.Start();
        //取出當日資料
        var dataList = conn.Query<LOWDATA>(sql, new { _strTime, _endTime });
        stopwatch1.Stop();
        TimeSpan timeSpan1 = stopwatch1.Elapsed;
        Console.Write("從資料庫取出資料時間:" + timeSpan1);


        Stopwatch stopwatch2 = new Stopwatch();
        stopwatch2.Start();

        //取出資料庫最早的日期，當成遞迴跳出的節點
        sql = $"SELECT Min(TIME) FROM [AIOT].[dbo].[Machine_Data]";
        //遞迴跳出的時間點
        var lastTimeData = Convert.ToDateTime(conn.QueryFirstOrDefault<string>(sql));
        //取得前一次工單編號
        sql = "SELECT [DeviceName],[WorKCode] FROM [AIOT].[dbo].[TempWork]";
        var tempBeforeWorkCodeList = conn.Query<TempWorkCode>(sql);

        //LowData整理
        #region
        foreach (var item in dataList)
        {
            //判斷設備名稱，如果沒有該設備名稱，則建立一個Temp
            var oldData = tempLowDatas.FirstOrDefault(x => x.DeviceName == item.DeviceName);
            if (oldData == null)
            {
                var temp = await createTemp(item);
                tempLowDatas.Add(temp);
            }
            //取出該筆設備Temp資料
            oldData = tempLowDatas.FirstOrDefault(x => x.DeviceName == item.DeviceName);

            //如果當天沒有輸入過WKC的話，WorKCodec會是NULL要撈取前一天的WKC，工單編號只取第一台
            if (string.IsNullOrEmpty(oldData.WorkCode))
            {
                var tempWorkCode = tempBeforeWorkCodeList.FirstOrDefault(x => x.DeviceName == oldData.DeviceName)?.WorkCode;

                oldData.WorkCode = tempWorkCode == null ? findWKC(conn, sql, item.DeviceName, _strTime, _endTime, lastTimeData) : tempWorkCode;
                //oldData.WorkCode = findWKC(conn, sql, item.DeviceName, _strTime, _endTime, lastTimeData);
            }

            //判斷不是關機的狀態關機斷
            if (item.Quality == "Good")
            {
                //判斷品檢模式
                if (item.Name.ToUpper().Contains("_QIM_"))
                {
                    if (item.Value == "1")
                    {
                        if (string.IsNullOrEmpty(oldData.StartTime))
                        {
                            oldData.StartTime = item.Time;
                        }
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        countOutQIMTime(tempNonWorkData, completeNonWorkDataS, item);
                        //品檢模式抓取6S時間防止在自動運行時重複抓取
                        count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);

                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork == null)
                        {
                            //2024 / 4 / 1新增
                            //移除缺料模式，如果是先自動運行在進入品檢模式會經過缺料，所以移除若沒有經過缺料代表沒有自動運行直接進入品檢
                            if (oldData.DMISuperMode == true)
                            {
                                var tempDMI = tempNonWorkData.FirstOrDefault(x => x.Name.ToUpper().Contains("_DMI_") && x.DeviceName == item.DeviceName);
                                if (tempDMI != null)
                                {
                                    tempNonWorkData.Remove(tempDMI);
                                }
                            }
                            //無開機工時紀錄
                            NonWork nonWork = new NonWork();
                            nonWork.WorkCode = oldData.WorkCode;
                            nonWork.DeviceOrder = item.DeviceOrder;
                            nonWork.DeviceName = item.DeviceName;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            //nonWork.Factory = item.Factory;
                            nonWork.Line = item.ProductLine;
                            nonWork.Activation = item.Activation;
                            nonWork.Throughput = item.Throughput;
                            nonWork.Defective = item.Defective;
                            nonWork.Exception = item.Exception;
                            nonWork.Description = item.Description;
                            nonWork.StartTime = item.Time;
                            nonWork.Name = item.Name;
                            nonWork.Date = _Date;
                            tempNonWorkData.Add(nonWork);
                            oldData.State = "品檢";
                            //看板計時總時間
                            oldData.ModelStartTime = item.Time;
                        }
                    }
                    else
                    {
                        //無開機工時紀錄
                        //判斷是否有異常資料，例如連續0
                        var tempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name == item.Name && x.DeviceName == item.DeviceName);
                        //2024/04/26增加防止資料庫資料時間有問題
                        var DMItempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name.Contains("_DMI_") && x.DeviceName == item.DeviceName);
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                            //2024 / 4 / 1新增
                            //退出品檢模式後，跳出機故、缺料的視窗，為防止操作人員沒有按按鈕的動作
                            if (DMItempNonWork == null)
                            {
                                NonWork nonWorkStopQIMTime = new NonWork();
                                nonWorkStopQIMTime.WorkCode = oldData.WorkCode;
                                nonWorkStopQIMTime.DeviceOrder = item.DeviceOrder;
                                nonWorkStopQIMTime.DeviceName = item.DeviceName;
                                nonWorkStopQIMTime.Activation = item.Activation;
                                nonWorkStopQIMTime.Throughput = item.Throughput;
                                nonWorkStopQIMTime.Defective = item.Defective;
                                nonWorkStopQIMTime.Exception = item.Exception;
                                nonWorkStopQIMTime.Product = item.Product;
                                nonWorkStopQIMTime.Item = item.Item;
                                //nonWorkStopQIMTime.Factory = item.Factory;
                                nonWorkStopQIMTime.Line = item.ProductLine;
                                nonWorkStopQIMTime.StartTime = item.Time;
                                nonWorkStopQIMTime.EndTime = item.Time;
                                nonWorkStopQIMTime.Name = "StopQTime";
                                nonWorkStopQIMTime.Description = "退出品檢到缺料停機時間";
                                nonWorkStopQIMTime.Date = _Date;
                                tempNonWorkData.Add(nonWorkStopQIMTime);
                                oldData.State = "未自動運行";
                                //看板計時總時間
                                oldData.ModelStartTime = item.Time;
                            }
                        }
                    }
                    oldData.QIMSuperMode = item.Value == "1" ? true : false;
                }
                //判斷缺料停機改善
                //進入缺料停機改善前，機器一定是處於自動運轉的狀態，當關閉缺料停機改善時，自動運行狀態一定開的
                //缺料停機改善1的時候自動運轉一定為0，缺料停機改善0的時候自動運轉一定為1
                if (item.Name.ToUpper().Contains("_DMI_"))
                {

                    if (item.Value == "1")
                    {
                        if (string.IsNullOrEmpty(oldData.StartTime))
                        {
                            oldData.StartTime = item.Time;
                        }
                        ///退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        countOutQIMTime(tempNonWorkData, completeNonWorkDataS, item);
                        //2024/4/22新增
                        //直接進入品檢模式抓取6S時間防止在自動運行時重複抓取
                        count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                        var tempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name == item.Name && x.DeviceName == item.DeviceName);
                        if (tempNonWork == null)
                        {
                            //無開機工時紀錄
                            NonWork nonWork = new NonWork();
                            nonWork.WorkCode = oldData.WorkCode;
                            nonWork.DeviceOrder = item.DeviceOrder;
                            nonWork.DeviceName = item.DeviceName;
                            nonWork.Activation = item.Activation;
                            nonWork.Throughput = item.Throughput;
                            nonWork.Defective = item.Defective;
                            nonWork.Exception = item.Exception;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            //nonWork.Factory = item.Factory;
                            nonWork.Line = item.ProductLine;
                            nonWork.Description = item.Description;
                            nonWork.StartTime = item.Time;
                            nonWork.Name = item.Name;
                            nonWork.Date = _Date;
                            tempNonWorkData.Add(nonWork);
                            oldData.State = "缺料停機改善";

                            //看板計時總時間
                            oldData.ModelStartTime = item.Time;
                        }
                    }
                    else
                    {
                        //無開機工時紀錄
                        var tempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name == item.Name && x.DeviceName == item.DeviceName);
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                            //看板計時總時間
                            oldData.ModelStartTime = null;
                        }

                    }
                    oldData.DMISuperMode = item.Value == "1" ? true : false;
                }
                //判斷機台故障維修
                if (item.Name.ToUpper().Contains("_MTC_"))
                {
                    if (item.Value == "1")
                    {
                        if (string.IsNullOrEmpty(oldData.StartTime))
                        {
                            oldData.StartTime = item.Time;
                        }
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        countOutQIMTime(tempNonWorkData, completeNonWorkDataS, item);
                        count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                        var tempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name == item.Name && x.DeviceName == item.DeviceName);
                        if (tempNonWork == null)
                        {
                            // 設備損失工時
                            NonWork nonWork = new NonWork();
                            nonWork.WorkCode = oldData.WorkCode;
                            nonWork.DeviceOrder = item.DeviceOrder;
                            nonWork.DeviceName = item.DeviceName;
                            nonWork.Activation = item.Activation;
                            nonWork.Throughput = item.Throughput;
                            nonWork.Defective = item.Defective;
                            nonWork.Exception = item.Exception;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            //nonWork.Factory = item.Factory;
                            nonWork.Line = item.ProductLine;
                            nonWork.Description = item.Description;
                            nonWork.StartTime = item.Time;
                            nonWork.Name = item.Name;
                            nonWork.Date = _Date;
                            tempNonWorkData.Add(nonWork);
                            oldData.State = "機台故障維修";

                            //看板計時總時間
                            oldData.ModelStartTime = item.Time;
                        }
                    }
                    else
                    {
                        // 設備損失工時
                        var tempNonWork = tempNonWorkData.FirstOrDefault(x => x.Name == item.Name && x.DeviceName == item.DeviceName);
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                            //看板計時總時間
                            oldData.ModelStartTime = null;
                        }

                    }
                    oldData.MTCSuperMode = item.Value == "1" ? true : false;
                }
                //品檢模式不計算產量、不良總數、CR不良、CCD1不良等等
                if (oldData.QIMSuperMode != true)
                {
                    //判斷不良品分類
                    if (item.Name.ToUpper().Contains("_NGI_"))
                    {
                        //oldData.RunState == "1"拿掉在自動運行狀態下才計算
                        string type = "NGI";
                        CountERRAndPath(item, oldData, type, _Date, dailyERRDatas);

                    }
                    //判斷工單數量
                    if (item.Name.ToUpper().Contains("_WKS_"))
                    {
                        if (item.Value != "0")
                        {
                            oldData.Sum += 1;
                        }
                    }
                    //判斷不良總量
                    if (item.Name.ToUpper().Contains("_NGS_"))
                    {
                        oldData.NGSum = Convert.ToInt32(item.Value);
                    }
                }
                //判斷是否處於無開機工時模式(品檢、缺料停機)
                if (oldData.QIMSuperMode != true && oldData.DMISuperMode != true && oldData.MTCSuperMode != true)
                {
                    //判斷機台啟動
                    if (item.Name.ToUpper().Contains("_RUN_"))
                    {
                        if (item.Value == "1")
                        {
                            // 儲存機台狀況
                            oldData.RunState = item.Value;
                            //看板計時總時間
                            oldData.ModelStartTime = null;

                            //如果之前停過機台、在開機時清除EndTime會造成下列問題
                            //產線可能只跑白班，但隔天八點前會暖機，所以沒有啟動就不清除EndTime
                            oldData.Quality = item.Quality;
                            oldData.EndTime = null;
                            oldData.RunStartTime = item.Time;
                            oldData.State = "自動運行中";
                            oldData.SumTime = null;
                            //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                            countOutQIMTime(tempNonWorkData, completeNonWorkDataS, item);
                            if (string.IsNullOrEmpty(oldData.StartTime))
                            {
                                oldData.StartTime = item.Time;
                            }
                            //計算6S
                            count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);

                            //計算RunSumStopTime如果沒有RunEndTime(沒有停機時間)就不進行RunSumStopTime的計算，計算停機10分鐘以上的機器
                            if (!string.IsNullOrEmpty(oldData.RunEndTime))
                            {
                                //開機停機時間收集
                                MachineStop machineStop = new MachineStop();
                                machineStop.DeviceName = item.DeviceName;
                                machineStop.WorkCode = oldData.WorkCode;
                                machineStop.Date = _Date;

                                var ts = (Convert.ToDateTime(oldData.RunStartTime) - Convert.ToDateTime(oldData.RunEndTime)).TotalMinutes;
                                //大於10分鐘以上才算做停機 ==> 稼動率參數使用
                                if (ts > 10)
                                {
                                    // 開機停機時間收集
                                    //十分鐘以下
                                    machineStop.StartTime = oldData.RunEndTime;
                                    machineStop.EndTime = oldData.RunStartTime;
                                    machineStop.SumTime = Math.Round((Convert.ToDateTime(machineStop.EndTime) - Convert.ToDateTime(machineStop.StartTime)).TotalMinutes, 2);
                                    oldData.StopTenUP.Add(machineStop);
                                    //大於10分鐘以上加入RunSumStopTime後面計算稼動率
                                    oldData.RunSumStopTime += ts;
                                }
                                else
                                {
                                    //開機停機時間收集
                                    //取出上次的停機時間
                                    //十分鐘以下
                                    machineStop.StartTime = oldData.RunEndTime;
                                    machineStop.EndTime = item.Time;
                                    machineStop.SumTime = Math.Round((Convert.ToDateTime(machineStop.EndTime) - Convert.ToDateTime(machineStop.StartTime)).TotalMinutes, 2);
                                    oldData.StopTenDown.Add(machineStop);
                                }

                                //新增2024 - 02 - 29計算是上次停機後清除RunEndTime
                                oldData.RunEndTime = "";
                            }
                        }
                        else
                        {
                            //如果RunState == null 代表工單未執行自動運行、缺料、機故
                            if (oldData.RunState != null)
                            {
                                oldData.RunEndTime = item.Time;
                                oldData.State = "未自動運行";
                                // 儲存機台狀況
                                oldData.RunState = item.Value;
                                //看板計時總時間
                                oldData.ModelStartTime = item.Time;
                            }
                        }
                    }
                    //判斷PAT錯誤，開機五分鐘內的訊息不紀錄
                    if (item.Name.ToUpper().Contains("_PAT_"))
                    {
                        if (item.Value == "1" && oldData.RunState == "1")
                        {
                            //開機五分鐘後才紀錄異常碼
                            if (Convert.ToDateTime(oldData.RunStartTime).AddMinutes(5) < Convert.ToDateTime(item.Time))
                            {
                                string type = "PAT";
                                CountERRAndPath(item, oldData, type, _Date, dailyERRDatas);
                            }
                        }
                    }
                    //判斷ERR錯誤，開機五分鐘內的訊息不紀錄
                    if (item.Name.ToUpper().Contains("_ERR_"))
                    {
                        if (item.Value == "1" && oldData.RunState == "1")
                        {
                            //開機五分鐘後才紀錄異常碼
                            if (Convert.ToDateTime(oldData.RunStartTime).AddMinutes(5) < Convert.ToDateTime(item.Time))
                            {
                                //新增2024 / 3 / 18 收到異常碼後兩秒的不紀錄
                                if (string.IsNullOrEmpty(oldData.ERRCountTime))
                                {
                                    oldData.ERRCountTime = item.Time;
                                    string type = "ERR";
                                    CountERRAndPath(item, oldData, type, _Date, dailyERRDatas);
                                }
                                else if (Convert.ToDateTime(oldData.ERRCountTime).AddSeconds(2) < Convert.ToDateTime(item.Time))
                                {
                                    oldData.ERRCountTime = item.Time;
                                    string type = "ERR";
                                    CountERRAndPath(item, oldData, type, _Date, dailyERRDatas);
                                }
                            }
                        }
                    }

                }
                //判斷工單編號
                if (item.Name.ToUpper().Contains("_WKC_") && !string.IsNullOrEmpty(item.Value) && item.Value.Length == 13)
                {
                    if (item.Value.Trim() != oldData.WorkCode)
                    {
                        //工單號進來的時間等於舊工單號結束的時間
                        //2024/04/19 註解 關機時間判斷錯誤
                        oldData.EndTime = item.Time;
                        oldData.Quality = "Bad";
                        oldData.SumTime = (Convert.ToDateTime(oldData.EndTime) - Convert.ToDateTime(oldData.StartTime)).TotalMinutes;

                        oldData.State = "完成";
                        oldData.ModelStartTime = null;

                        //移除TempNonTime的資料
                        tempNonWorkData.RemoveAll(x => x.DeviceName == item.DeviceName && x.WorkCode == oldData.WorkCode);

                        //更換工單                    
                        completeLowDatas.Add(oldData);
                        tempLowDatas.Remove(oldData);
                        var newData = await createTemp(item);
                        newData.WorkCode = item.Value.TrimStart().TrimEnd();
                        tempLowDatas.Add(newData);
                    }
                }
            }
            //關機寫入關機時間，寫入關機狀態，關機流程: 停機 >> (缺料停機改善、機台故障維修)>> 關機，關機前把缺料停機改善及機台故障維修清除
            else
            {
                //關機後開機會送新的WorkCode所以相當於換工單
                //只有第一筆會進行新增建工單，
                if (!string.IsNullOrEmpty(oldData.StartTime))
                {
                    oldData.State = "完成";
                    oldData.EndTime = item.Time;
                    oldData.Quality = item.Quality;
                    oldData.RunState = "0";
                    oldData.SumTime = (Convert.ToDateTime(oldData.EndTime) - Convert.ToDateTime(oldData.StartTime)).TotalMinutes;
                    oldData.ModelStartTime = null;

                    completeLowDatas.Add(oldData);
                    var newData = await createTemp(item);
                    newData.WorkCode = oldData.WorkCode;
                    tempLowDatas.Remove(oldData);
                    tempLowDatas.Add(newData);
                }
                //關機時把缺料、機故時間結算
                var tempNonWork = tempNonWorkData.Where(x => x.DeviceName == item.DeviceName).Select(x => x).ToList();
                if (tempNonWork.Count > 0)
                {
                    foreach (var temp in tempNonWork)
                    {
                        temp.EndTime = item.Time;
                        temp.SumTime = (Convert.ToDateTime(temp.EndTime) - Convert.ToDateTime(temp.StartTime)).TotalMinutes;
                        completeNonWorkDataS.Add(temp);
                        tempNonWorkData.Remove(temp);
                    }
                }
            }
        }
        stopwatch2.Stop();
        TimeSpan timeSpan2 = stopwatch2.Elapsed;
        Console.Write("資料分析時間:" + timeSpan2);
        #endregion
        //匯總資料
        #region
        completeLowDatas.AddRange(tempLowDatas);
        completeNonWorkDataS.AddRange(tempNonWorkData);


        //2024/03/06新增
        //判斷產線有沒有執行，如果沒有從completeLowDatas移除
        //沒有RunStartTime等於沒有自動運行
        //completeLowDatas.RemoveAll(x => string.IsNullOrEmpty(x.RunStartTime) || Convert.ToInt32(x.Sum) <= 1);
        completeLowDatas.RemoveAll(x => string.IsNullOrEmpty(x.RunStartTime));
        #endregion
        //當日有資料才進行分析
        if (completeLowDatas.Count > 0)
        {
            inputData(out sql, _strTime, _endTime, _Date, conn, completeLowDatas, completeNonWorkDataS, dailyERRDatas);
        }
    }
}
static void countOutQIMTime(List<NonWork> tempNonWorkData, List<NonWork> completeNonWorkDataS, LOWDATA item)
{
    var tempNonWorkStopQTime = tempNonWorkData.Where(x => x.Name == "StopQTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
    if (tempNonWorkStopQTime != null)
    {
        tempNonWorkStopQTime.EndTime = item.Time;
        tempNonWorkStopQTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQTime.StartTime)).TotalMinutes;
        tempNonWorkData.Remove(tempNonWorkStopQTime);
        completeNonWorkDataS.Add(tempNonWorkStopQTime);
    }
}
static void CountERRAndPath(LOWDATA item, TempData? oldData, string type, string Date, List<DailyERRData> dailyERRDatas)
{
    //錯誤訊息計數
    var names = item.Name.ToUpper().Split('_');
    //取出寄存器名稱
    var patName = names[4];
    var key = patName + "_" + item.Description;

    var data = dailyERRDatas.Where(x => x.Name == key && x.DeviceName == item.DeviceName && x.WorkCode == oldData.WorkCode).FirstOrDefault();
    if (data != null)
    {
        //只有臨停才進行時間紀錄
        if (type == "NGI")
        {
            data.Count = Convert.ToInt32(item.Value);
        }
        else
        {
            if (type == "ERR")
            {
                data.Time += Convert.ToDateTime(item.Time).ToString("yyyy-MM-dd HH:mm:ss") + "\n";
            }
            data.Count++;
        }


    }
    else
    {
        DailyERRData dailyERRData = new DailyERRData();
        dailyERRData.Name = key;
        //只有臨停才進行時間紀錄
        if (type == "ERR")
        {
            dailyERRData.Time = Convert.ToDateTime(item.Time).ToString("yyyy-MM-dd HH:mm:ss") + "\n";
        }
        dailyERRData.Date = Date;
        dailyERRData.Line = item.ProductLine;
        dailyERRData.WorkCode = oldData.WorkCode;
        dailyERRData.DeviceName = item.DeviceName;
        dailyERRData.Type = type;
        dailyERRData.Count = type == "NGI" ? Convert.ToInt32(item.Value) : 1;
        dailyERRDatas.Add(dailyERRData);

    }
}
static void SendEMail(string Conetext, bool check)
{
    //寄件者帳號密碼
    string senderEmail = "aiot@dip.net.cn";
    string password = "dipf2aiot";

    //收件者


    MailMessage mail = new MailMessage();
    mail.From = new MailAddress(senderEmail);
    mail.Subject = "AIOT系統通知信";
    mail.Body = Conetext;
    //可以加入多個收件者
    if (check)
    {
        mail.To.Add("ibukiboy@dip.com.tw");
        mail.To.Add("why1@dip.net.cn");
        mail.To.Add("luobing@dip.net.cn");
        mail.CC.Add("gary.tsai@dip.com.tw");
    }
    else
    {
        mail.CC.Add("gary.tsai@dip.com.tw");
    }

    SmtpClient smtpServer = new SmtpClient("mail.dip.net.cn");
    smtpServer.Port = 25;
    smtpServer.Credentials = new NetworkCredential(senderEmail, password);
    smtpServer.EnableSsl = false;

    smtpServer.Send(mail);

}
static void count6STime(string _strTime, string _Date, List<NonWork> completeNonWorkDataS, LOWDATA item, TempData? oldData)
{
    var check = completeNonWorkDataS.Where(x => x.Name == "6S" && x.DeviceName == item.DeviceName).FirstOrDefault();
    if (check == null)
    {
        //只有第一筆資料WorkCode是NULL
        oldData.StartTime = item.Time;
        //無開機工時紀錄 6S
        NonWork nonWork6S = new NonWork();
        nonWork6S.WorkCode = oldData.WorkCode;
        nonWork6S.DeviceOrder = item.DeviceOrder;
        nonWork6S.DeviceName = item.DeviceName;
        nonWork6S.Activation = item.Activation;
        nonWork6S.Throughput = item.Throughput;
        nonWork6S.Defective = item.Defective;
        nonWork6S.Exception = item.Exception;
        nonWork6S.Product = item.Product;
        nonWork6S.Item = item.Item;
        //nonWork6S.Factory = item.Factory;
        nonWork6S.Line = item.ProductLine;
        nonWork6S.StartTime = _strTime;
        nonWork6S.EndTime = item.Time;
        nonWork6S.Name = "6S";
        nonWork6S.Description = "6S";
        nonWork6S.Date = _Date;
        nonWork6S.SumTime = (Convert.ToDateTime(nonWork6S.EndTime) - Convert.ToDateTime(nonWork6S.StartTime)).TotalMinutes;
        completeNonWorkDataS.Add(nonWork6S);
    }
}

static void UpdateNonTime(List<NonWork> tempNonWorkData, List<NonWork> completeNonWorkDataS, LOWDATA item, TempData? oldData, NonWork? tempNonWork)
{
    oldData.RunEndTime = "";
    oldData.RunStartTime = item.Time;
    oldData.RunState = "1";

    tempNonWork.EndTime = item.Time;
    tempNonWork.SumTime = (Convert.ToDateTime(tempNonWork.EndTime) - Convert.ToDateTime(tempNonWork.StartTime)).TotalMinutes;
    completeNonWorkDataS.Add(tempNonWork);
    tempNonWorkData.Remove(tempNonWork);
    oldData.State = "自動運行中";
}



