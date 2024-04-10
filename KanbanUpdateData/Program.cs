
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
    tempData.ERRAndPATCountList = new List<DailyERRData>();
    tempData.StopTenUP = new List<MachineStop>();
    tempData.StopTenDown = new List<MachineStop>();
    var Date = item.Time.Split(' ');
    tempData.Line = item.ProductLine;
    tempData.Factory = item.Factory;
    tempData.Alloted = item.Alloted;
    tempData.State = "未運行";
    tempData.Folor = item.Folor;
    tempData.Model = item.Model;
    tempData.Date = Date[0];
    tempData.DeviceName = item.DeviceName;
    tempData.Item = item.Item;
    tempData.Product = item.Product;
    tempData.Activation = item.Activation;
    tempData.Throughput = item.Throughput;
    tempData.Defective = item.Defective;
    tempData.Exception = item.Exception;
    tempData.QIMSuperMode = false;
    tempData.DMISuperMode = false;
    return tempData;
}
string findWKC(SqlConnection con, string sql, string DeviceName, string strTime, string endTime, DateTime lastTimeData)
{
    strTime = Convert.ToDateTime(strTime).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss");
    endTime = Convert.ToDateTime(endTime).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss");
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
string findPNO(SqlConnection con, string sql, string DeviceName, string strTime, string endTime, DateTime lastTimeData)
{
    strTime = Convert.ToDateTime(strTime).ToString("yyyy-MM-dd HH:mm:ss");
    endTime = Convert.ToDateTime(endTime).ToString("yyyy-MM-dd HH:mm:ss");
    if (Convert.ToDateTime(strTime) < lastTimeData)
    {
        return "找不到P數";
    }
    else
    {
        sql = $"SELECT TOP(1) VALUE FROM [AIOT].[dbo].[Machine_Data] where NAME like '%PNO%' and time between '{strTime}' and '{endTime}' AND Quality <> 'Bad' AND DEVICENAME like '%{DeviceName}%' AND VALUE <> '' AND VALUE IS NOT NULL  order by TIME desc";
        var data = con.QueryFirstOrDefault<string>(sql);
        if (data != null)
        {
            return data.TrimStart().TrimEnd();
        }
        else
        {
            strTime = Convert.ToDateTime(strTime).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss");
            endTime = Convert.ToDateTime(endTime).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss");
            data = findPNO(con, sql, DeviceName, strTime, endTime, lastTimeData);
            return data.TrimStart().TrimEnd();
        }
    }
}
void inputData(out string sql, string _strTime, string _endTime, string _Date, SqlConnection conn, List<TempData> completeLowDatas, List<NonWork> completeNonWorkDataS)
{
    //取出資料庫最早的日期，當成遞迴跳出的節點
    sql = $"SELECT Min(TIME) FROM [AIOT].[dbo].[Machine_Data]";
    //遞迴跳出的時間點
    var lastTimeData = Convert.ToDateTime(conn.QueryFirstOrDefault<string>(sql));

    foreach (var item in completeLowDatas)
    {
        //如果當天沒有輸入過WKC的話，WorKCodec會是NULL要撈取前一天的WKC，工單編號只取第一台
        if (string.IsNullOrEmpty(item.WorkCode) && item.Activation == true)
        {
            item.WorkCode = findWKC(conn, sql, item.DeviceName, _strTime, _endTime, lastTimeData);
        }
        else if (string.IsNullOrEmpty(item.WorkCode) && item.Activation == false)
        {
            item.WorkCode = "";
        }
        //  EndTime如果是NULL的話代表執行24小時
        if (string.IsNullOrEmpty(item.EndTime))
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

    //試做工單排除統計
    //"150R-"
    var machineDataList = completeLowDatas.Where(x => !x.WorkCode.Contains("150R-")).GroupBy(x =>
    new { x.Item, x.Product, x.Line, x.Factory, x.DeviceName, x.Alloted, x.Folor, x.Activation, x.Throughput, x.Defective, x.Exception }).Select(y => new
    {
        y.Key.Alloted,
        y.Key.Folor,
        y.Key.Line,
        y.Key.Item,
        y.Key.Product,
        y.Key.Factory,
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

    //取出瓶警機當日工作時間最長的WorkCode
    var tempStdList = completeLowDatas.Where(x => x.Activation == true && !x.WorkCode.Contains("150R-")).GroupBy(x => new { x.Factory, x.Item, x.Product, x.Line, x.WorkCode }).Select(y => new
    {
        y.Key.Factory,
        y.Key.WorkCode,
        y.Key.Product,
        y.Key.Line,
        y.Key.Item,
        Model = completeLowDatas.Where(x => x.Activation == true && !x.WorkCode.Contains("150R-")).Select(x => x.Model).FirstOrDefault(),
        AllTime = (Convert.ToDateTime(y.Max(z => z.EndTime)) - Convert.ToDateTime(y.Min(z => z.StartTime))).TotalMinutes
    }).GroupBy(x => new { x.Factory, x.Product, x.Line, x.Item }).Select(y => y.OrderByDescending(z => z.AllTime).First()).ToList();
    //工單 搜尋 ERP資料庫找到對應的料件，料件對應LOCAL DB 找出品名、PCS
    using (OracleConnection oracleConnection = new OracleConnection(produndtConnectionString))
    {

        foreach (var item1 in tempStdList)
        {
            var partsql = $"select sfb05 as Part_No from dipf2.sfb_file,dipf2.ima_file where sfb05 =ima01 and sfb87='Y' and sfb01 = '{item1.WorkCode}'";
            var part = oracleConnection.Query<string>(partsql).FirstOrDefault();
            if (part != null)
            {
                var PCS = stdPerformanceList.Where(x => x.Part_No == part.ToString() && x.Model == item1.Model).Select(x => x.PCS).FirstOrDefault();
                var data = new TempStd();
                data.WorkCode = item1.WorkCode;
                data.PCS = PCS;
                data.AllTime = item1.AllTime.ToString();
                data.Item = item1.Item;
                data.Line = item1.Line;
                data.Product = item1.Product;
                data.Factory = item1.Factory;
                data.Product_Name = stdPerformanceList.Where(x => x.Part_No == part.ToString()).Select(x => x.Product_Name).FirstOrDefault();
                pcsList.Add(data);
            }
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
    //無開機工時: 6S、缺料開機、品檢模式匯入資料庫
    #region
    var dailyNonWorkDatas = completeNonWorkDataS.GroupBy(x => new { x.Item, x.Product, x.Line, x.Alloted, x.Folor, x.Factory, x.DeviceName, x.Name, x.Activation, x.Throughput, x.Defective, x.Exception }).Select(x => new
    {
        x.Key.DeviceName,
        x.Key.Name,
        x.Key.Activation,
        x.Key.Throughput,
        x.Key.Defective,
        x.Key.Exception,
        x.Key.Alloted,
        x.Key.Folor,
        x.Key.Product,
        x.Key.Line,
        x.Key.Factory,
        x.Key.Item,
        _Date,
        SumCount = x.Sum(y => Convert.ToDouble(y.SumTime)),
    });
    //臨停缺料List
    var dailyERRDatas = new List<DailyERRData>();
    foreach (var item in completeLowDatas)
    {
        dailyERRDatas.AddRange(item.ERRAndPATCountList);
    }
    //取出最後一台機的臨停
    var stopAndMachineList = from a in machineDataList
                             join b in dailyERRDatas on a.DeviceName equals b.DeviceName
                             where b.Type == "ERR"
                             select new
                             {
                                 a.Item,
                                 a.Factory,
                                 a.Product,
                                 a.Line,
                                 b.Count,
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
    var Line_MachineDailyData = machineDataList.Where(x => x.Exception != true).GroupBy(x => new { x.Factory, x.Line, x.Product, x.Item, x.Alloted, x.Folor }).Select(y =>
    {
        //預計投入工時(Throughput)
        var ETC = Math.Round((Convert.ToDateTime(y.Where(x => x.Throughput == true).Select(x => x.MaxTime).FirstOrDefault()) - Convert.ToDateTime(_strTime)).TotalHours, 2);
        // 6S(Throughput)
        var Non6sTime = dailyNonWorkDatas.Where(x => x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("6S")).Select(x => x.SumCount).FirstOrDefault() / 60;
        //缺料停機(Throughput)
        var NonDMITime = dailyNonWorkDatas.Where(x => x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("DMI")).Select(x => x.SumCount).FirstOrDefault() / 60;
        // 品檢模式(Throughput)
        var NonQIMTime = dailyNonWorkDatas.Where(x => x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("QIM")).Select(x => x.SumCount).FirstOrDefault() / 60;
        //退出品檢模式到缺料停機或機台改善之前時間(Throughput)
        var NonStopQIMTime = dailyNonWorkDatas.Where(x => x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("StopQIMTime")).Select(x => x.SumCount).FirstOrDefault() / 60;
        //無開機工時
        var NonTime = Non6sTime + NonDMITime + NonQIMTime + NonStopQIMTime;
        //設備損失工時已排除(Exception)//自動運行狀態下臨停10分鐘以上
        var StopRunTime = y.Where(x => x.Throughput == true).Select(x => x.StopRunTime).FirstOrDefault(0.0) / 60;
        //機台故障維修//人員操作機故障時間
        var MTCTime = dailyNonWorkDatas.Where(x => x.Throughput == true && x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Product == y.Key.Product && x.Item == y.Key.Item && x.Name.Contains("MTC")).Select(x => x.SumCount).FirstOrDefault() / 60;
        StopRunTime = StopRunTime + MTCTime;

        var PT = Math.Round(ETC - NonTime, 2);
        var ACT = Math.Round(PT - StopRunTime, 2);
        //顯示當日更換的所有品名
        var Product_Name = string.Join(", ", Product_NameList.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).GroupBy(x => x.Product_Name).Select(x => x.Key));
        var SC = Convert.ToDouble(pcsList.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).Select(y => y.PCS).FirstOrDefault());
        //實際產量(Throughput)
        var AO = y.Where(x => x.Throughput == true).Select(x => x.Sum).FirstOrDefault(0.0);
        var YieIdAO = y.Where(x => x.Defective == true && x.Throughput != true).Select(x => x.Sum).FirstOrDefault(0.0);
        var countYieId = y.Where(x => x.Defective == true).Select(x => x.NGS).DefaultIfEmpty(0.0).Sum();
        var Performance = Math.Round(((AO / SC) / ACT) * 100, 2).ToString();
        //(測試機 + 全檢(不良數)) / 測試產出量 * 100
        var YieId = (100 - Math.Round((countYieId) / YieIdAO * 100, 2)).ToString();
        var Availability = Math.Round(((ACT / PT) * 100), 2).ToString();
        var OEE = Math.Round((Convert.ToDouble(Performance) / 100) * (Convert.ToDouble(YieId) / 100) * (Convert.ToDouble(Availability) / 100) * 100, 2).ToString();
        //目前產線有幾台機器
        var MachineCount = completeLowDatas.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).GroupBy(x => x.DeviceName).Count();
        //全部機台的臨停
        var StopCount = stopAndMachineList.Where(x => x.Line == y.Key.Line && x.Factory == y.Key.Factory && x.Item == y.Key.Item && x.Product == y.Key.Product).Sum(x => x.Count);
        //平均臨停
        var AVGStopCount = Math.Round(Convert.ToDouble(StopCount / ACT / MachineCount), 2).ToString();
        //最後一台機目前運行狀態
        var State = y.Where(x => x.Throughput == true).Select(x => x.State).FirstOrDefault("未運行");
        return new
        {

            Product_Name,
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
            AVGStopCount
        };
    });
    //先清空TABLE再新增資料
    sql = "DELETE [AIOT].[dbo].[KanBan_Line_MachineData] ";
    conn.Execute(sql);
    sql = @" insert into [AIOT].[dbo].[KanBan_Line_MachineData]
                            Values(@Factory,@Item,@Product,@State,@Alloted,@Folor,@Product_Name,@Line,@Date,@MCT,@SC,@ETC,@PT,@ACT,@ACTH,@AO,@CAPU,@ADR,@Performance,@YieId,@Availability,@OEE,@NonTime,@StopRunTime,@AVGStopCount)";
    //執行資料庫
    conn.Execute(sql, Line_MachineDailyData);
    //無開機資料匯入資料庫
    //sql = @"insert into [AIOT].[dbo].[Line_MachineNonTime]
    //                        Values(@DeviceName,@Description,@Name,@Date,@StartTime,@EndTime,@SumTime)";
    ////執行資料庫
    ////conn.Execute(sql, completeNonWorkDataS);
    //#endregion
    ////統整錯誤訊息、10分鐘以上與10分鐘以下開機時間紀錄
    //#region
    //var dailyERRDatas = new List<DailyERRData>();
    //var TenUpDatas = new List<MachineStop>();
    //var TenDownDatas = new List<MachineStop>();
    //foreach (var item in completeLowDatas)
    //{
    //    TenUpDatas.AddRange(item.StopTenUP);
    //    TenDownDatas.AddRange(item.StopTenDown);
    //    dailyERRDatas.AddRange(item.ERRAndPATCountList);
    //}
    ////錯誤訊息DailyERRDatas匯入資料庫
    //sql = @"insert into [AIOT].[dbo].[Line_MachineERRData]
    //                        Values(@DeviceName,@Date,@Time,@Type,@Name,@Count)";
    ////執行資料庫
    ////conn.Execute(sql, dailyERRDatas);

    ////10分鐘以上開機時間紀錄匯入資料庫
    //sql = "INSERT INTO [AIOT].[dbo].[Line_Machine_StopTenUp] VALUES(@DeviceName,@Date,@StartTime,@EndTime,@SumTime)";
    ////執行資料庫
    ////conn.Execute(sql, TenUpDatas);

    ////10分鐘以下開機時間紀錄匯入資料庫
    //sql = "INSERT INTO [AIOT].[dbo].[Line_Machine_StopTenDown] VALUES(@DeviceName,@Date,@StartTime,@EndTime,@SumTime)";
    ////執行資料庫
    ////conn.Execute(sql, TenDownDatas);
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
    var _Date = Convert.ToDateTime(_strTime).ToString("yyyy-MM-dd");
    //取出當天的LowData
    string sql = @"SELECT F.Factory, F.Item,F.Product,F.Alloted,F.Folor,F.Model,F.ProductLine,F.Activation,F.Throughput,F.Defective,F.Exception,MD.DeviceName,MD.NAME,MD.QUALITY,MD.TIME,MD.VALUE,MD.Description ";
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
    using (SqlConnection conn = new SqlConnection(connectionString))
    {
        var dataList = conn.Query<LOWDATA>(sql);
        //LowData整理
        #region
        foreach (var item in dataList)
        {
            var oldData = tempLowDatas.FirstOrDefault(x => x.DeviceName == item.DeviceName, null);
            //判斷設備名稱，如果沒有該設備名稱，則建立一個Temp
            if (oldData == null)
            {
                var temp = createTemp(item);
                tempLowDatas.Add(temp);
            }
            //取出該筆設備Temp資料
            oldData = tempLowDatas.First(x => x.DeviceName == item.DeviceName);
            //判斷不是關機的狀態關機斷
            if (item.Quality != "Bad")
            {
                // 這邊不洗掉EndTime時間，自動運行後才會洗掉(Run)
                if (oldData.Quality == "Bad")
                {
                    oldData.Quality = item.Quality;
                    oldData.State = "未運行";
                }
                //判斷品檢模式
                if (item.Name.ToUpper().Contains("QIM"))
                {
                    if (item.Value == "1")
                    {
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
                            var check = completeNonWorkDataS.Where(x => x.Description.Equals("6S") && x.DeviceName == item.DeviceName).FirstOrDefault();
                            if (check == null)
                            {
                                //只有第一筆資料WorkCode是NULL
                                oldData.StartTime = item.Time;
                                //無開機工時紀錄 6S
                                NonWork nonWork6S = new NonWork();
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

                            //無開機工時紀錄
                            NonWork nonWork = new NonWork();
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
                        if (tempNonWork != null)
                        {
                            //重製機台開機及關機
                            oldData.RunEndTime = "";
                            oldData.RunStartTime = item.Time;
                            oldData.RunState = "1";
                            //無開機工時紀錄
                            tempNonWork.EndTime = item.Time;
                            tempNonWork.SumTime = (Convert.ToDateTime(tempNonWork.EndTime) - Convert.ToDateTime(tempNonWork.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWork);
                            completeNonWorkDataS.Add(tempNonWork);
                            //2024 / 4 / 1新增
                            //退出品檢模式後，跳出機故、缺料的視窗，為防止操作人員沒有按按鈕的動作

                            NonWork nonWorkStopQIMTime = new NonWork();
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
                            nonWorkStopQIMTime.Name = "StopQIMTime";
                            nonWorkStopQIMTime.Description = "退出品檢到缺料停機時間";
                            nonWorkStopQIMTime.Date = _Date;
                            tempNonWorkData.Add(nonWorkStopQIMTime);
                            oldData.State = "未運行";
                        }
                    }
                    oldData.QIMSuperMode = item.Value == "1" ? true : false;
                }
                //判斷缺料停機改善
                //進入缺料停機改善前，機器一定是處於自動運轉的狀態，當關閉缺料停機改善時，自動運行狀態一定開的
                //缺料停機改善1的時候自動運轉一定為0，缺料停機改善0的時候自動運轉一定為1
                if (item.Name.ToUpper().Contains("DMI"))
                {
                    if (item.Value == "1")
                    {
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        var tempNonWorkStopQIMTime = tempNonWorkData.Where(x => x.Name == "StopQIMTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWorkStopQIMTime != null)
                        {
                            tempNonWorkStopQIMTime.EndTime = item.Time;
                            tempNonWorkStopQIMTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQIMTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQIMTime.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWorkStopQIMTime);
                            completeNonWorkDataS.Add(tempNonWorkStopQIMTime);
                        }
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork == null)
                        {
                            //無開機工時紀錄
                            NonWork nonWork = new NonWork();
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
                            //缺料進入前會寫入RunEndtime出去後會啟動自動運轉會計算10鐘以上，RunEndTime清除這樣就不會計算缺料的時間
                            oldData.RunEndTime = "";
                            oldData.RunStartTime = item.Time;
                            oldData.RunState = "1";

                            tempNonWork.EndTime = item.Time;
                            tempNonWork.SumTime = (Convert.ToDateTime(tempNonWork.EndTime) - Convert.ToDateTime(tempNonWork.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWork);
                            completeNonWorkDataS.Add(tempNonWork);
                            oldData.State = "未運行";
                        }
                    }
                    oldData.DMISuperMode = item.Value == "1" ? true : false;
                }
                //判斷機台故障維修
                if (item.Name.ToUpper().Contains("MTC"))
                {
                    if (item.Value == "1")
                    {
                        //退出品檢模式時如果沒有案的話就會計算中間的空白時間
                        var tempNonWorkStopQIMTime = tempNonWorkData.Where(x => x.Name == "StopQIMTime" && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWorkStopQIMTime != null)
                        {
                            tempNonWorkStopQIMTime.EndTime = item.Time;
                            tempNonWorkStopQIMTime.SumTime = (Convert.ToDateTime(tempNonWorkStopQIMTime.EndTime) - Convert.ToDateTime(tempNonWorkStopQIMTime.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWorkStopQIMTime);
                            completeNonWorkDataS.Add(tempNonWorkStopQIMTime);
                        }
                        var tempNonWork = tempNonWorkData.Where(x => x.Name == item.Name && x.DeviceName == item.DeviceName).FirstOrDefault();
                        if (tempNonWork == null)
                        {
                            // 設備損失工時
                            NonWork nonWork = new NonWork();
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
                            oldData.RunEndTime = "";
                            oldData.RunStartTime = item.Time;
                            oldData.RunState = "1";

                            tempNonWork.EndTime = item.Time;
                            tempNonWork.SumTime = (Convert.ToDateTime(tempNonWork.EndTime) - Convert.ToDateTime(tempNonWork.StartTime)).TotalMinutes;
                            tempNonWorkData.Remove(tempNonWork);
                            completeNonWorkDataS.Add(tempNonWork);
                            oldData.State = "未運行";
                        }
                    }
                    oldData.MTCSuperMode = item.Value == "1" ? true : false;
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
                        machineStop.Date = _Date;
                        // 儲存機台狀況
                        oldData.RunState = item.Value;
                        if (item.Value == "1")
                        {
                            //機器最早啟動時間，其他時間因為換工單時會直接填入工單編號，所以只有第一筆會沒有工單編號。

                            var check = completeNonWorkDataS.Where(x => x.Description.Equals("6S") && x.DeviceName == item.DeviceName).FirstOrDefault();
                            if (check == null)
                            {
                                // 只要是已經建立過的TMEP一定會有工單號，沒有工單號代表是程式最初建立的TEMP
                                oldData.StartTime = item.Time;
                                // 無開機工時紀錄 6S
                                NonWork nonWork = new NonWork();
                                nonWork.DeviceName = item.DeviceName;
                                nonWork.Activation = item.Activation;
                                nonWork.Throughput = item.Throughput;
                                nonWork.Defective = item.Defective;
                                nonWork.Exception = item.Exception;
                                nonWork.Product = item.Product;
                                nonWork.Alloted = item.Alloted;
                                nonWork.Folor = item.Folor;
                                nonWork.Item = item.Item;
                                nonWork.Factory = item.Factory;
                                nonWork.Line = item.ProductLine;
                                nonWork.StartTime = _strTime;
                                nonWork.EndTime = item.Time;
                                nonWork.Name = "6S";
                                nonWork.Description = "6S";
                                nonWork.Date = _Date;
                                nonWork.SumTime = (Convert.ToDateTime(nonWork.EndTime) - Convert.ToDateTime(nonWork.StartTime)).TotalMinutes;
                                completeNonWorkDataS.Add(nonWork);
                            }

                            //如果之前有停機過，又啟動機器在把EndTime清出(產線可能只跑白班，但隔天八點前會暖機，所以沒有啟動就不清除EndTime
                            oldData.EndTime = null;
                            oldData.RunStartTime = item.Time;
                            oldData.State = "自動運行中";
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
                            oldData.State = "未運行";
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
                                CountERRAndPat(item, oldData, type, _Date);
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
                                    CountERRAndPat(item, oldData, type, _Date);
                                }
                                else if (Convert.ToDateTime(oldData.ERRCountTime).AddSeconds(2) < Convert.ToDateTime(item.Time))
                                {
                                    oldData.ERRCountTime = item.Time;
                                    string type = "ERR";
                                    CountERRAndPat(item, oldData, type, _Date);
                                }
                            }
                        }
                    }
                    //判斷工單數量
                    if (item.Name.ToUpper().Contains("WKS"))
                    {
                        if (item.Value != "0")
                        {
                            oldData.Sum = (Convert.ToDouble(oldData.Sum) + 1).ToString();
                        }
                    }
                    //判斷不良總量
                    if (item.Name.ToUpper().Contains("NGS"))
                    {
                        if (item.Value != "0")
                        {
                            oldData.NGSum = (Convert.ToDouble(oldData.NGSum) + 1).ToString();
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
                        oldData.EndTime = item.Time;
                        completeLowDatas.Add(oldData);
                        tempLowDatas.Remove(oldData);
                        var newData = createTemp(item);
                        newData.WorkCode = item.Value.TrimStart().TrimEnd();
                        tempLowDatas.Add(newData);
                    }
                }
            }
            //關機寫入關機時間，寫入關機狀態，關機流程: 停機 >> (缺料停機改善、機台故障維修)>> 關機，關機前把缺料停機改善及機台故障維修清除
            else
            {
                //移除無開機工時裡的Temp
                var tempNonWork = tempNonWorkData.Where(x => x.DeviceName == item.DeviceName).ToList();
                if (tempNonWork.Count > 0)
                {
                    foreach (var data in tempNonWork)
                    {
                        tempNonWorkData.Remove(data);
                    }
                }
                oldData.State = "關機";
                oldData.EndTime = item.Time;
                oldData.Quality = item.Quality;
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
        var tempNoRunList = completeLowDatas.Where(x => string.IsNullOrEmpty(x.RunStartTime)).Select(x => x).ToList();
        if (tempNoRunList.Count > 0)
        {
            foreach (var data in tempNoRunList)
            {
                completeLowDatas.Remove(data);
            }
        }
        #endregion
        //當日有資料才進行分析
        if (completeLowDatas.Count > 0)
        {
            inputData(out sql, _strTime, _endTime, _Date, conn, completeLowDatas, completeNonWorkDataS);
        }
    }
}

void CountERRAndPat(LOWDATA item, TempData? oldData, string type, string Date)
{
    //錯誤訊息計數
    var names = item.Name.ToUpper().Split('_');
    //取出寄存器名稱
    var patName = names[4];
    var key = patName + "_" + item.Description;

    var data = oldData.ERRAndPATCountList.Where(x => x.Name == key).FirstOrDefault();
    if (data != null)
    {
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
        if (type == "ERR")
        {
            dailyERRData.Time = Convert.ToDateTime(item.Time).ToString("yyyy-MM-dd HH:mm:ss") + "\n";
        }
        dailyERRData.Date = Date;
        dailyERRData.Line = item.ProductLine;
        dailyERRData.DeviceName = item.DeviceName;
        dailyERRData.Type = type;
        dailyERRData.Count = 1;
        oldData.ERRAndPATCountList.Add(dailyERRData);

    }
}

//void SendEMail(string Conetext)
//{
//    //寄件者帳號密碼
//    string senderEmail = "aiot@dip.net.cn";
//    string password = "dipf2aiot";

//    //收件者


//    MailMessage mail = new MailMessage();
//    mail.From = new MailAddress(senderEmail);
//    mail.Subject = "AIOT系統通知信";
//    mail.Body = Conetext;
//    //可以加入多個收件者
//    mail.To.Add("ibukiboy@dip.com.tw");
//    mail.To.Add("why1@dip.net.cn");
//    mail.To.Add("luobing@dip.net.cn");
//    mail.CC.Add("gary.tsai@dip.com.tw");

//    SmtpClient smtpServer = new SmtpClient("mail.dip.net.cn");
//    smtpServer.Port = 25;
//    smtpServer.Credentials = new NetworkCredential(senderEmail, password);
//    smtpServer.EnableSsl = false;

//    smtpServer.Send(mail);
//    Console.WriteLine("邮件发送成功！");

//}



