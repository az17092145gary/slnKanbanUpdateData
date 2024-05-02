
using System.Net.Mail;
using System.Net;
using KanbanUpdateData.Model;
using System.Data.SqlClient;
using Dapper;
using Oracle.ManagedDataAccess.Client;



string connectionString = "Data Source= 192.168.0.82;Initial Catalog=AIOT;Persist Security Info=True;User ID=sa;Password=P@ssw0rd;Encrypt=True;TrustServerCertificate=True";
string produndtConnectionString = "User ID=ds;Password=ds;Data Source=192.168.160.207:1521/topprod";


executeMethod();



TempData createTemp(LOWDATA item)
{


    TempData tempData = new TempData();
    tempData.StopTenUP = new List<MachineStop>();
    tempData.StopTenDown = new List<MachineStop>();
    var Date = item.Time.Split(' ');
    tempData.Line = item.ProductLine;
    tempData.Factory = item.Factory;
    tempData.Alloted = item.Alloted;
    tempData.State = "未運行";
    tempData.Folor = item.Folor;
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
    tempData.RunStartTime = item.Time;
    tempData.StartTime = item.Time;
    tempData.QIMSuperMode = false;
    tempData.DMISuperMode = false;
    return tempData;
}
string findWKC(SqlConnection con, string sql, string DeviceName, string strTime, string endTime, DateTime lastTimeData)
{
    strTime = Convert.ToDateTime(strTime).AddDays(-1).ToString("yyyy-MM-dd 08:00:00");
    endTime = Convert.ToDateTime(endTime).ToString("yyyy-MM-dd 08:00:00");
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
            endTime = Convert.ToDateTime(endTime).AddDays(-1).ToString("yyyy-MM-dd 08:00:00");
            data = findWKC(con, sql, DeviceName, strTime, endTime, lastTimeData);
            return data.TrimStart().TrimEnd();
        }
    }
}
void inputData(out string sql, string _strTime, string _endTime, string _Date, SqlConnection conn, List<TempData> completeLowDatas, List<NonWork> completeNonWorkDataS, List<DailyERRData> dailyERRDatas)
{

    foreach (var item in completeLowDatas)
    {

        //如果Quality不等於Bad代表運行24小時或者並無關機
        if (item.Quality != "Bad")
        {
            item.EndTime = _endTime;
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



    //日報表不用排除150R試做工單，看板系統需要排除
    var machineDataList = completeLowDatas.Where(x => !x.WorkCode.Contains("150R-")).GroupBy(x =>
    new { x.WorkCode, x.DeviceOrder, x.Item, x.Product, x.Line, x.Factory, x.DeviceName, x.Alloted, x.Folor, x.Activation, x.Throughput, x.Defective, x.Exception }).Select(y => new
    {
        y.Key.WorkCode,
        y.Key.Alloted,
        y.Key.Folor,
        y.Key.Line,
        y.Key.Item,
        y.Key.Product,
        y.Key.Factory,
        y.Key.DeviceOrder,
        y.Key.DeviceName,
        y.Key.Activation,
        y.Key.Throughput,
        y.Key.Defective,
        y.Key.Exception,
        State = y.OrderByDescending(x => x.EndTime).FirstOrDefault()?.State,
        MinTime = y.Min(z => z.StartTime) ?? null,
        MaxTime = y.Max(z => z.EndTime) ?? null,
        Sum = y.Sum(z => Convert.ToDouble(z.Sum)),
        NGS = y.Sum(z => Convert.ToDouble(z.NGSum)),
        StopRunTime = y.Sum(z => Convert.ToDouble(z.RunSumStopTime))
    }).ToList();
    //資料庫取出標準產能
    #region

    sql = $"SELECT * FROM [AIOT].[dbo].[Standard_Production_Efficiency_Benchmark]";
    var stdPerformanceList = conn.Query<StdPerformance>(sql).ToList();
    //標準產能(依照當日執行時間最長的工單抓取標準產能)
    var pcsList = new List<TempStd>();
    //全部品名
    var Product_NameList = new List<TempStd>();

    var tempStdList = completeLowDatas.Where(x => x.Activation == true && !x.WorkCode.Contains("150R-")).GroupBy(x => new { x.Factory, x.Item, x.Product, x.Line, x.WorkCode }).Select(y => new
    {
        y.Key.Factory,
        y.Key.WorkCode,
        y.Key.Product,
        y.Key.Line,
        y.Key.Item,
        Model = completeLowDatas.Where(x => x.Activation == true && !x.WorkCode.Contains("150R-")).Select(x => x.Model).FirstOrDefault(),
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
            var PCS = stdPerformanceList.Where(x => x.Part_No == part.ToString() && x.Model == item1.Model).Select(x => x.PCS).FirstOrDefault();
            if (PCS == null)
            {
                var tempList = machineDataList.Where(x => x.Line == item1.Line && x.Item == item1.Item && x.Product == item1.Product).Select(x => x).ToList();
                machineDataList = machineDataList.Except(tempList).ToList();
                emailContext += _Date + $"{item1.WorkCode}該工單對應不到標準產能。";
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
                data.Factory = item1.Factory;
                data.Product_Name = stdPerformanceList.Where(x => x.Part_No == part.ToString()).Select(x => x.Product_Name).FirstOrDefault();
                pcsList.Add(data);
            }
        }
        //對應不到標準產能的工單通知生產人員
        if (!string.IsNullOrEmpty(emailContext))
        {
            SendEMail(emailContext);
        }
        //新增: 2024 / 03 / 18 取出當日瓶警機的全部品名
        foreach (var item2 in completeLowDatas)
        {
            if (item2.Activation == true)
            {
                var partsql = $"select sfb05 as Part_No from dipf2.sfb_file,dipf2.ima_file where sfb05 =ima01 and sfb87='Y' and sfb01 = '{item2.WorkCode}'";
                var part = oracleConnection.Query<string>(partsql).FirstOrDefault();
                if (part != null)
                {
                    var data = new TempStd();
                    data.WorkCode = item2.WorkCode;
                    data.Item = item2.Item;
                    data.Line = item2.Line;
                    data.Product = item2.Product;
                    data.Factory = item2.Factory;
                    data.Product_Name = stdPerformanceList.Where(x => x.Part_No == part.ToString()).Select(x => x.Product_Name).FirstOrDefault();
                    Product_NameList.Add(data);
                }
            }
        }

    }
    #endregion
    #region
    //排除PAT(目前沒有在看)
    dailyERRDatas = dailyERRDatas.Where(x => x.Type != "PAT").Select(y => y).ToList();
    //取出工單臨停
    var stopAndMachineList = from a in machineDataList
                             join b in dailyERRDatas on a.DeviceName equals b.DeviceName
                             where b.Type == "ERR"
                             group a by new { a.Item, a.Factory, a.Product, a.Line, b.Count, a.WorkCode } into g
                             select new
                             {
                                 Item = g.Key.Item,
                                 Factory = g.Key.Factory,
                                 Product = g.Key.Product,
                                 Line = g.Key.Line,
                                 Count = g.Key.Count,
                                 WorkCode = g.Key.WorkCode,
                                 TotalCount = g.Count()
                             };

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
    var Line_MachineDailyData = machineDataList.GroupBy(x => new { x.WorkCode, x.Factory, x.Line, x.Product, x.Item, x.Alloted, x.Folor }).Select(y =>
    {
        var MachineCount = completeLowDatas.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).GroupBy(x => x.DeviceName).Count();
        //工單開始時間
        var WKCEndTime = Convert.ToDateTime(y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.MaxTime).FirstOrDefault());
        //工單結束時間
        var WKCStartTime = Convert.ToDateTime(y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.MinTime).FirstOrDefault());
        //關機時間
        var CloseTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("close")).Sum(x => x.SumTime) / 60;
        //預計投入工時(Throughput)
        var ETC = Math.Round((WKCEndTime - WKCStartTime).TotalHours, 2);
        // 6S(Throughput)
        //var Non6sTime = dailyNonWorkDatas.Where(x =>x.WorkCode == y.Key.WorkCode && x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("6S")).Select(x => x.SumCount).FirstOrDefault() / 60;
        //缺料停機(Throughput)
        var NonDMITime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("DMI") && Convert.ToDateTime(x.StartTime) >= WKCStartTime).Sum(x => x.SumTime) / 60;
        // 品檢模式(Throughput)
        var NonQIMTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("QIM") && Convert.ToDateTime(x.StartTime) >= WKCStartTime).Sum(x => x.SumTime) / 60;

        //退出品檢模式到缺料停機或機台改善之前時間(Throughput)
        var NonStopQTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("StopQTime") && Convert.ToDateTime(x.StartTime) >= WKCStartTime).Sum(x => x.SumTime) / 60;
        //2024/04/25
        //換品名時間
        var NonChangeProductName = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("ChangeProductName")).Sum(x => x.SumTime) / 60;
        //無開機工時
        var NonTime = NonDMITime + NonQIMTime + NonStopQTime + NonChangeProductName;
        //設備損失工時已排除(Exception)//自動運行狀態下臨停10分鐘以上
        var StopRunTime = y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.StopRunTime).FirstOrDefault(0.0) / 60;
        //機台故障維修//人員操作機故障時間
        var MTCTime = completeNonWorkDataS.Where(x => x.WorkCode == y.Key.WorkCode && Convert.ToInt32(x.DeviceOrder) == MachineCount && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("MTC") && Convert.ToDateTime(x.StartTime) >= WKCStartTime).Sum(x => x.SumTime) / 60;
        StopRunTime = StopRunTime + MTCTime;

        var PT = Math.Round(ETC - NonTime, 2);
        var ACT = Math.Round(PT - StopRunTime, 2);
        //顯示當日更換的所有品名
        var Product_Name = string.Join(", ", Product_NameList.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).GroupBy(x => x.Product_Name).Select(x => x.Key));
        var SC = Convert.ToDouble(pcsList.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).Select(y => y.PCS).FirstOrDefault());
        //實際產量(Throughput)
        var AO = y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.Sum).FirstOrDefault(0.0);
        var lastYieIDAO = y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.NGS).FirstOrDefault(0.0);
        AO = AO + lastYieIDAO;
        var YieIdAO = y.Where(x => x.Defective == true && x.Throughput != true).Select(x => x.Sum).FirstOrDefault(0.0);
        var AllNGS = y.Where(x => x.Defective == true).Select(x => x.NGS).DefaultIfEmpty(0.0).Sum();
        var Performance = Math.Round(((AO / SC) / ACT) * 100, 2).ToString();
        //(測試機 + 全檢(不良數)) / 測試產出量 * 100
        var YieId = (100 - Math.Round((AllNGS) / YieIdAO * 100, 2)).ToString();
        var Availability = Math.Round(((ACT / PT) * 100), 2).ToString();
        var OEE = Math.Round((Convert.ToDouble(Performance) / 100) * (Convert.ToDouble(YieId) / 100) * (Convert.ToDouble(Availability) / 100) * 100, 2).ToString();
        //工單總臨停
        var StopCount = stopAndMachineList.Where(x => x.WorkCode == y.Key.WorkCode && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).Sum(x => x.Count);
        //平均臨停
        var AVGStopCount = Math.Round(Convert.ToDouble(StopCount / ACT / MachineCount), 2).ToString();
        //最後一台機目前運行狀態
        var State = y.Where(x => Convert.ToInt32(x.DeviceOrder) == MachineCount).Select(x => x.State).FirstOrDefault("未運行");
        return new
        {
            NonChangeProductName,
            NonDMITime,
            NonQIMTime,
            NonStopQTime,
            MachineCount,
            WKCStartTime,
            WKCEndTime,
            Product_Name,
            y.Key.WorkCode,
            y.Key.Factory,
            y.Key.Item,
            y.Key.Product,
            State,
            y.Key.Alloted,
            y.Key.Folor,
            y.Key.Line,
            Date = _Date,
            MCT = "24",
            SC,
            ETC,
            PT,
            ACT,
            ACTH = Math.Round((AO / SC), 2),
            AO,
            CAPU = Math.Round(((ETC / 24) * 100), 2),
            ADR = Math.Round((PT / ETC) * 100, 2),
            Performance,
            YieId,
            Availability,
            OEE,
            NonTime = Math.Round(NonTime, 2),
            StopRunTime = Math.Round(StopRunTime, 2),
            AVGStopCount,
            AllNGS
        };
    });

    //2024/04/19 排除異常數據
    Line_MachineDailyData = Line_MachineDailyData.Where(x => x.AO > 0 && Convert.ToDouble(x.Performance) > 0).Select(x => x).ToList();
    //刪除看板系統紀錄
    sql = " DELETE [AIOT].[dbo].[KanBan_Line_MachineData] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_Machine_StopTenDown] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_Machine_StopTenUp] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_MachineERRData] ";
    sql += " DELETE [AIOT].[dbo].[KanBan_MachineNonTime] ";
    conn.Execute(sql);
    sql = @" insert into [AIOT].[dbo].[KanBan_Line_MachineData]
                            Values(@Factory,@Item,@Product,@State,@Alloted,@Folor,@WorkCode,@Product_Name,@Line,@Date,@MCT,@SC,@ETC,@PT,@ACT,@ACTH,@AO,@CAPU,@ADR,@Performance,@YieId,@Availability,@OEE,@NonTime,@StopRunTime,@AVGStopCount,@AllNGS)";
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

void executeMethod()
{
    //開始時間
    var _strTime = DateTime.Today.Add(new TimeSpan(8, 0, 0)).ToString("yyyy-MM-dd HH:mm:ss");
    //結束時間
    var _endTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    //今日日期
    //_strTime = "2024-04-29 08:00:00";
    //_endTime = "2024-04-30 08:00:00";
    var _Date = Convert.ToDateTime(_strTime).ToString("yyyy-MM-dd");
    //取出當天的LowData
    string sql = @"SELECT F.Factory, F.Item,F.Product,F.Alloted,F.Folor,F.Model,F.DeviceOrder,F.ProductLine,F.Activation,F.Throughput,F.Defective,F.Exception,MD.DeviceName,MD.NAME,MD.QUALITY,MD.TIME,MD.VALUE,MD.Description ";
    sql += $" FROM(select * FROM[AIOT].[dbo].[Machine_Data] WHERE TIME BETWEEN '{_strTime}' AND '{_endTime}') AS MD";
    sql += " LEFT JOIN [AIOT].[dbo].[Factory]as F ON F.[IODviceName] = MD.[DeviceName] ";
    sql += " ORDER BY TIME";

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
        //取出當日資料
        var dataList = conn.Query<LOWDATA>(sql);

        //取出資料庫最早的日期，當成遞迴跳出的節點
        sql = $"SELECT Min(TIME) FROM [AIOT].[dbo].[Machine_Data]";
        //遞迴跳出的時間點
        var lastTimeData = Convert.ToDateTime(conn.QueryFirstOrDefault<string>(sql));
        //LowData整理
        #region
        foreach (var item in dataList)
        {
            //判斷設備名稱，如果沒有該設備名稱，則建立一個Temp
            var oldData = tempLowDatas.Where(x => x.DeviceName == item.DeviceName).FirstOrDefault();
            if (oldData == null)
            {
                var temp = createTemp(item);
                tempLowDatas.Add(temp);
            }
            //取出該筆設備Temp資料
            oldData = tempLowDatas.Where(x => x.DeviceName == item.DeviceName).FirstOrDefault();
            //如果當天沒有輸入過WKC的話，WorKCodec會是NULL要撈取前一天的WKC，工單編號只取第一台
            if (string.IsNullOrEmpty(oldData.WorkCode))
            {
                oldData.WorkCode = findWKC(conn, sql, item.DeviceName, _strTime, _endTime, lastTimeData);
            }

            //判斷不是關機的狀態關機斷
            if (item.Quality != "Bad")
            {
                //如果有關機時間就計算總時數
                var closeNonWork = tempNonWorkData.Where(x => x.DeviceName == item.DeviceName && x.Name == "close").Select(x => x).FirstOrDefault();
                if (closeNonWork != null)
                {
                    closeNonWork.EndTime = item.Time;
                    closeNonWork.SumTime = (Convert.ToDateTime(closeNonWork.EndTime) - Convert.ToDateTime(closeNonWork.StartTime)).TotalMinutes;
                    tempNonWorkData.Remove(closeNonWork);
                    completeNonWorkDataS.Add(closeNonWork);
                }
                //判斷品檢模式
                if (item.Name.ToUpper().Contains("QIM"))
                {
                    if (item.Value == "1")
                    {
                        oldData.EndTime = null;
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        var tempNonWorkStopQTime = tempNonWorkData.Where(x => x.Name == "StopQTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWorkStopQTime != null)
                        {
                            tempNonWorkStopQTime.EndTime = item.Time;
                            tempNonWorkStopQTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQTime.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWorkStopQTime);
                            completeNonWorkDataS.Add(tempNonWorkStopQTime);
                        }

                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork == null)
                        {
                            //2024 / 4 / 1新增
                            //移除缺料模式，如果是先自動運行在進入品檢模式會經過缺料，所以移除若沒有經過缺料代表沒有自動運行直接進入品檢
                            if (oldData.DMISuperMode == true)
                            {
                                var tempDMI = tempNonWorkData.Where(x => x.Name.ToUpper().Contains("DMI") && x.DeviceName == item.DeviceName).FirstOrDefault();
                                if (tempDMI != null)
                                {
                                    tempNonWorkData.Remove(tempDMI);
                                }
                            }
                            //2024 / 4 / 1新增
                            //直接進入品檢模式抓取6S時間防止在自動運行時重複抓取
                            count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                            //無開機工時紀錄
                            NonWork nonWork = new NonWork();
                            nonWork.WorkCode = oldData.WorkCode;
                            nonWork.DeviceOrder = item.DeviceOrder;
                            nonWork.DeviceName = item.DeviceName;
                            nonWork.Alloted = item.Alloted;
                            nonWork.Folor = item.Folor;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            nonWork.Factory = item.Factory;
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
                        }
                    }
                    else
                    {
                        //無開機工時紀錄
                        //判斷是否有異常資料，例如連續0
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        //2024/04/26增加防止資料庫資料時間有問題
                        var DMItempNonWork = tempNonWorkData.Where(x => x.Name.Contains("DMI") && x.DeviceName == item.DeviceName).FirstOrDefault();
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
                                nonWorkStopQIMTime.Alloted = item.Alloted;
                                nonWorkStopQIMTime.Folor = item.Folor;
                                nonWorkStopQIMTime.Item = item.Item;
                                nonWorkStopQIMTime.Factory = item.Factory;
                                nonWorkStopQIMTime.Line = item.ProductLine;
                                nonWorkStopQIMTime.StartTime = item.Time;
                                nonWorkStopQIMTime.EndTime = item.Time;
                                nonWorkStopQIMTime.Name = "StopQTime";
                                nonWorkStopQIMTime.Description = "退出品檢到缺料停機時間";
                                nonWorkStopQIMTime.Date = _Date;
                                tempNonWorkData.Add(nonWorkStopQIMTime);
                                oldData.State = "未運行";
                            }
                        }
                    }
                    oldData.QIMSuperMode = item.Value == "1" ? true : false;
                }
                //判斷缺料停機改善
                //進入缺料停機改善前，機器一定是處於自動運轉的狀態，當關閉缺料停機改善時，自動運行狀態一定開的
                //缺料停機改善1的時候自動運轉一定為0，缺料停機改善0的時候自動運轉一定為1
                if (item.Name.ToUpper().Contains("DMI"))
                {

                    oldData.EndTime = null;
                    if (item.Value == "1")
                    {
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        var tempNonWorkStopQIMTime = tempNonWorkData.Where(x => x.Name == "StopQTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWorkStopQIMTime != null)
                        {
                            tempNonWorkStopQIMTime.EndTime = item.Time;
                            tempNonWorkStopQIMTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQIMTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQIMTime.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWorkStopQIMTime);
                            completeNonWorkDataS.Add(tempNonWorkStopQIMTime);
                        }
                        //2024/4/22新增
                        //直接進入品檢模式抓取6S時間防止在自動運行時重複抓取
                        count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
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
                            nonWork.Alloted = item.Alloted;
                            nonWork.Folor = item.Folor;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            nonWork.Factory = item.Factory;
                            nonWork.Line = item.ProductLine;
                            nonWork.Description = item.Description;
                            nonWork.StartTime = item.Time;
                            nonWork.Name = item.Name;
                            nonWork.Date = _Date;
                            tempNonWorkData.Add(nonWork);
                            oldData.State = "缺料停機改善";
                        }
                    }
                    else
                    {
                        //無開機工時紀錄
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                        }
                        tempNonWork = tempNonWorkData.Where(x => x.Name == "ChangeProductName" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                        }
                        //2024-04-19新增:缺料停機結束=自動運行啟動
                        if (string.IsNullOrEmpty(oldData.StartTime))
                        {
                            oldData.StartTime = item.Time;
                        }
                    }
                    oldData.DMISuperMode = item.Value == "1" ? true : false;
                }
                //判斷機台故障維修
                if (item.Name.ToUpper().Contains("MTC"))
                {
                    if (item.Value == "1")
                    {
                        oldData.EndTime = null;

                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        var tempNonWorkStopQIMTime = tempNonWorkData.Where(x => x.Name == "StopQTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWorkStopQIMTime != null)
                        {
                            tempNonWorkStopQIMTime.EndTime = item.Time;
                            tempNonWorkStopQIMTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQIMTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQIMTime.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWorkStopQIMTime);
                            completeNonWorkDataS.Add(tempNonWorkStopQIMTime);
                        }
                        count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
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
                            nonWork.Alloted = item.Alloted;
                            nonWork.Folor = item.Folor;
                            nonWork.Product = item.Product;
                            nonWork.Item = item.Item;
                            nonWork.Factory = item.Factory;
                            nonWork.Line = item.ProductLine;
                            nonWork.Description = item.Description;
                            nonWork.StartTime = item.Time;
                            nonWork.Name = item.Name;
                            nonWork.Date = _Date;
                            tempNonWorkData.Add(nonWork);
                            oldData.State = "機台故障維修";
                        }
                    }
                    else
                    {
                        // 設備損失工時
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                        }
                        tempNonWork = tempNonWorkData.Where(x => x.Name == "ChangeProductName" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork != null)
                        {
                            UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                        }
                        //2024-04-19新增:缺料停機結束=自動運行啟動
                        if (string.IsNullOrEmpty(oldData.StartTime))
                        {
                            oldData.StartTime = item.Time;
                        }
                    }
                    oldData.MTCSuperMode = item.Value == "1" ? true : false;
                }
                //品檢模式不計算產量、不良總數、CR不良、CCD1不良等等
                if (oldData.QIMSuperMode != true)
                {
                    //判斷不良品分類
                    if (item.Name.ToUpper().Contains("NGI"))
                    {
                        //oldData.RunState == "1"拿掉在自動運行狀態下才計算
                        if (item.Value != "0")
                        {
                            string type = "NGI";
                            CountERRAndPath(item, oldData, type, _Date, dailyERRDatas);
                        }
                    }
                    //判斷工單數量
                    if (item.Name.ToUpper().Contains("WKS"))
                    {
                        if (item.Value != "0")
                        {
                            oldData.Sum += 1;
                        }
                    }
                    //判斷不良總量
                    if (item.Name.ToUpper().Contains("NGS"))
                    {
                        if (item.Value != "0")
                        {
                            oldData.NGSum += 1;
                        }
                    }


                }
                //判斷是否處於無開機工時模式(品檢、缺料停機)
                if (oldData.QIMSuperMode != true && oldData.DMISuperMode != true && oldData.MTCSuperMode != true)
                {
                    //判斷機台啟動
                    if (item.Name.ToUpper().Contains("RUN"))
                    {

                        //開機停機時間收集
                        MachineStop machineStop = new MachineStop();
                        machineStop.DeviceName = item.DeviceName;
                        machineStop.WorkCode = oldData.WorkCode;
                        machineStop.Date = _Date;
                        // 儲存機台狀況
                        oldData.RunState = item.Value;
                        if (item.Value == "1")
                        {
                            //如果之前停過機台、在開機時清除EndTime會造成下列問題
                            //產線可能只跑白班，但隔天八點前會暖機，所以沒有啟動就不清除EndTime
                            oldData.Quality = item.Quality;
                            oldData.EndTime = null;
                            oldData.RunStartTime = item.Time;
                            oldData.State = "自動運行中";

                            //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                            var tempNonWorkStopQIMTime = tempNonWorkData.Where(x => x.Name == "StopQTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                            if (tempNonWorkStopQIMTime != null)
                            {
                                tempNonWorkStopQIMTime.EndTime = item.Time;
                                tempNonWorkStopQIMTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQIMTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQIMTime.StartTime)).TotalMinutes;
                                tempNonWorkData.Remove(tempNonWorkStopQIMTime);
                                completeNonWorkDataS.Add(tempNonWorkStopQIMTime);
                            }
                            //臨停時候換品名
                            var tempNonWork = tempNonWorkData.Where(x => x.Name == "ChangeProductName" && x.DeviceName == item.DeviceName).FirstOrDefault();
                            if (tempNonWork != null)
                            {
                                UpdateNonTime(tempNonWorkData, completeNonWorkDataS, item, oldData, tempNonWork);
                            }
                            //計算6S
                            count6STime(_strTime, _Date, completeNonWorkDataS, item, oldData);
                            //計算RunSumStopTime如果沒有RunEndTime(沒有停機時間)就不進行RunSumStopTime的計算，計算停機10分鐘以上的機器
                            if (!string.IsNullOrEmpty(oldData.RunEndTime))
                            {
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
                            oldData.RunEndTime = item.Time;
                            oldData.State = "未自動運行";
                        }
                    }
                    //判斷PAT錯誤，開機五分鐘內的訊息不紀錄
                    if (item.Name.ToUpper().Contains("PAT"))
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
                    if (item.Name.ToUpper().Contains("ERR"))
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
                if (item.Name.ToUpper().Contains("WKC") && !string.IsNullOrEmpty(item.Value) && item.Value.Length == 13 && item.Quality != "Bad")
                {

                    //2024 / 01 / 03增加check判斷值，確認再已有的資料中是否有相同工單編號
                    //可以避免前一天只有白班，隔天會在7: 50左右開機進行熱機送出重複的工單編號導致前一天的總時數不對
                    bool compCheck = false;
                    bool tempCheck = false;
                    foreach (var wkc in completeLowDatas)
                    {
                        compCheck = wkc.WorkCode == item.Value.TrimStart().TrimEnd() && wkc.DeviceName == item.DeviceName ? true : false;
                    }
                    foreach (var wkc in tempLowDatas)
                    {
                        tempCheck = wkc.WorkCode == item.Value.TrimStart().TrimEnd() && wkc.DeviceName == item.DeviceName ? true : false;
                    }
                    //item.Value的值 不等於 oldData代表有兩種狀況 1.oldData.WorkCode = null 或者 2.item.Value是新的工單號
                    if (oldData.WorkCode != item.Value.TrimStart().TrimEnd() && (!compCheck) && (!tempCheck))
                    {
                        //工單號進來的時間等於舊工單號結束的時間
                        //2024/04/19 註解 關機時間判斷錯誤
                        oldData.EndTime = item.Time;
                        oldData.Quality = "Bad";

                        var templist = tempNonWorkData.Where(x => x.DeviceName == item.DeviceName && x.WorkCode == oldData.WorkCode).OrderByDescending(x => x.StartTime).FirstOrDefault();

                        NonWork nonWorkChangeProductName;
                        if (templist != null)
                        {
                            //加上換品名時間
                            nonWorkChangeProductName = new NonWork();
                            nonWorkChangeProductName.WorkCode = oldData.WorkCode;
                            nonWorkChangeProductName.DeviceName = item.DeviceName;
                            nonWorkChangeProductName.DeviceOrder = item.DeviceOrder;
                            nonWorkChangeProductName.Activation = item.Activation;
                            nonWorkChangeProductName.Throughput = item.Throughput;
                            nonWorkChangeProductName.Defective = item.Defective;
                            nonWorkChangeProductName.Exception = item.Exception;
                            nonWorkChangeProductName.Product = item.Product;
                            nonWorkChangeProductName.Alloted = item.Alloted;
                            nonWorkChangeProductName.Folor = item.Folor;
                            nonWorkChangeProductName.Item = item.Item;
                            nonWorkChangeProductName.Factory = item.Factory;
                            nonWorkChangeProductName.Line = item.ProductLine;
                            nonWorkChangeProductName.StartTime = templist.StartTime;
                            nonWorkChangeProductName.EndTime = item.Time;
                            nonWorkChangeProductName.SumTime = (Convert.ToDateTime(nonWorkChangeProductName.EndTime) - Convert.ToDateTime(nonWorkChangeProductName.StartTime)).TotalMinutes;
                            nonWorkChangeProductName.Name = "ChangeProductName";
                            nonWorkChangeProductName.Description = "ChangeProductName";
                            nonWorkChangeProductName.Date = _Date;
                            completeNonWorkDataS.Add(nonWorkChangeProductName);
                        }
                        //移除TempNonTime的資料
                        tempNonWorkData.RemoveAll(x => x.DeviceName == item.DeviceName && x.WorkCode == oldData.WorkCode);
                        //更換工單
                        oldData.State = "完成";
                        completeLowDatas.Add(oldData);
                        tempLowDatas.Remove(oldData);
                        var newData = createTemp(item);
                        newData.WorkCode = item.Value.TrimStart().TrimEnd();
                        tempLowDatas.Add(newData);
                        //加上換品名時間
                        nonWorkChangeProductName = new NonWork();
                        nonWorkChangeProductName.WorkCode = newData.WorkCode;
                        nonWorkChangeProductName.DeviceOrder = item.DeviceOrder;
                        nonWorkChangeProductName.DeviceName = item.DeviceName;
                        nonWorkChangeProductName.Activation = item.Activation;
                        nonWorkChangeProductName.Throughput = item.Throughput;
                        nonWorkChangeProductName.Defective = item.Defective;
                        nonWorkChangeProductName.Exception = item.Exception;
                        nonWorkChangeProductName.Product = item.Product;
                        nonWorkChangeProductName.Alloted = item.Alloted;
                        nonWorkChangeProductName.Folor = item.Folor;
                        nonWorkChangeProductName.Item = item.Item;
                        nonWorkChangeProductName.Factory = item.Factory;
                        nonWorkChangeProductName.Line = item.ProductLine;
                        nonWorkChangeProductName.StartTime = item.Time;
                        nonWorkChangeProductName.Name = "ChangeProductName";
                        nonWorkChangeProductName.Description = "ChangeProductName";
                        nonWorkChangeProductName.Date = _Date;
                        tempNonWorkData.Add(nonWorkChangeProductName);
                    }
                }
            }
            //關機寫入關機時間，寫入關機狀態，關機流程: 停機 >> (缺料停機改善、機台故障維修)>> 關機，關機前把缺料停機改善及機台故障維修清除
            else
            {
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
                oldData.State = "關機";
                oldData.EndTime = item.Time;
                oldData.Quality = item.Quality;
                //2024/04/30 
                //關機時間紀錄
                var closeNonWork = tempNonWorkData.Where(x => x.DeviceName == item.DeviceName && x.Name == "close").Select(x => x).FirstOrDefault();
                if (closeNonWork == null)
                {
                    NonWork nonWorkClose = new NonWork();
                    nonWorkClose.WorkCode = oldData.WorkCode;
                    nonWorkClose.DeviceOrder = item.DeviceOrder;
                    nonWorkClose.DeviceName = item.DeviceName;
                    nonWorkClose.Activation = item.Activation;
                    nonWorkClose.Throughput = item.Throughput;
                    nonWorkClose.Defective = item.Defective;
                    nonWorkClose.Exception = item.Exception;
                    nonWorkClose.Product = item.Product;
                    nonWorkClose.Alloted = item.Alloted;
                    nonWorkClose.Folor = item.Folor;
                    nonWorkClose.Item = item.Item;
                    nonWorkClose.Factory = item.Factory;
                    nonWorkClose.Line = item.ProductLine;
                    nonWorkClose.StartTime = item.Time;
                    nonWorkClose.Name = "close";
                    nonWorkClose.Description = "close";
                    nonWorkClose.Date = _Date;
                    tempNonWorkData.Add(nonWorkClose);
                }
            }
        }

        #endregion
        //匯總資料
        #region
        completeLowDatas.AddRange(tempLowDatas);
        completeNonWorkDataS.AddRange(tempNonWorkData);


        //2024/03/06新增
        //判斷產線有沒有執行，如果沒有從completeLowDatas移除
        //沒有RunStartTime等於沒有自動運行
        completeLowDatas.RemoveAll(x => x.Sum < 5);
        #endregion
        //當日有資料才進行分析
        if (completeLowDatas.Count > 0)
        {
            inputData(out sql, _strTime, _endTime, _Date, conn, completeLowDatas, completeNonWorkDataS, dailyERRDatas);
        }
    }
}

void CountERRAndPath(LOWDATA item, TempData? oldData, string type, string Date, List<DailyERRData> dailyERRDatas)
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
        if (type == "ERR")
        {
            data.Time += Convert.ToDateTime(item.Time).ToString("yyyy-MM-dd HH:mm:ss") + "\n";
        }
        data.Count++;
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
        dailyERRData.Count = 1;
        dailyERRDatas.Add(dailyERRData);

    }
}

void count6STime(string _strTime, string _Date, List<NonWork> completeNonWorkDataS, LOWDATA item, TempData? oldData)
{
    var check = completeNonWorkDataS.Where(x => x.Description.Equals("6S") && x.DeviceName == item.DeviceName).FirstOrDefault();
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
        nonWork6S.Alloted = item.Alloted;
        nonWork6S.Folor = item.Folor;
        nonWork6S.Item = item.Item;
        nonWork6S.Factory = item.Factory;
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
    tempNonWorkData.Remove(tempNonWork);
    completeNonWorkDataS.Add(tempNonWork);
    oldData.State = "未運行";
}

void SendEMail(string Conetext)
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
    mail.To.Add("ibukiboy@dip.com.tw");
    mail.To.Add("why1@dip.net.cn");
    mail.To.Add("luobing@dip.net.cn");
    mail.CC.Add("gary.tsai@dip.com.tw");

    SmtpClient smtpServer = new SmtpClient("mail.dip.net.cn");
    smtpServer.Port = 25;
    smtpServer.Credentials = new NetworkCredential(senderEmail, password);
    smtpServer.EnableSsl = false;

    smtpServer.Send(mail);
    Console.WriteLine("邮件发送成功！");

}



