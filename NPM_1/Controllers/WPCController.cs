using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.InkML;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.Contracts;
using System.Diagnostics.SymbolStore;
using System.Linq;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.UI.WebControls;


namespace NPM_1.Controllers
{
    public class WPCController : Controller
    {
        public ActionResult WPC_KPI()    // (GET)
        {
            // Loads the initial view and populates dropdowns with provinces and corresponding cities.
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Provinces
                SqlCommand cmd1 = new SqlCommand("SELECT DISTINCT Province FROM SA_Province_Cont_MAP where Province!='Iran'", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                ViewBag.ProvinceList = dt1.AsEnumerable().Select(r => r[0].ToString()).ToList();



            }

            return View();
        }

        // Method of Query Execution with Output
        public DataTable Query_Execution_Table_Output(String Query)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            string Quary_String = Query;
            SqlCommand Quary_Command = new SqlCommand(Quary_String, conn);
            Quary_Command.CommandTimeout = 0;
            Quary_Command.ExecuteNonQuery();
            DataTable Output_Table = new DataTable();
            SqlDataAdapter dataAdapter_Quary_Command = new SqlDataAdapter(Quary_Command);
            dataAdapter_Quary_Command.Fill(Output_Table);
            return Output_Table;
        }


        public class TrafficFilter
        {
            public string Province { get; set; }
            public string Technology { get; set; }
            public List<string> Dates { get; set; }
            public string Interval { get; set; }
            public string KPI { get; set; }
            public string Threshold { get; set; }
            public string MinAvailability { get; set; }

            public string MinTraffic { get; set; }

        }

        public class FilterRequest
        {
            public List<string> Filters { get; set; }
            //public string AvailabilityThreshold { get; set; }
            //public string TrafficThreshold { get; set; }

        }

        public DataTable Data_Table_2G = new DataTable();
        public DataTable mergedTable2G = new DataTable();
        public DataTable Data_Table_3G = new DataTable();
        public DataTable mergedTable3G = new DataTable();
        public DataTable Data_Table_4G = new DataTable();
        public DataTable mergedTable4G = new DataTable();
        public DataTable mergedTable = new DataTable();

        [HttpPost]
        //public JsonResult FetchFilteredData(List<string> filters)
        public JsonResult FetchFilteredData([FromBody] FilterRequest request)
        {
            var parsedFilters = new List<TrafficFilter>();

            //string availabilityThreshold = request.AvailabilityThreshold;
            //string trafficThreshold = request.TrafficThreshold;

            // foreach (var filterText in filters)
            foreach (var filterText in request.Filters)

            {
                try
                {
                    var filter = new TrafficFilter();
                    var parts = filterText.Split(',');

                    var dateList = new List<string>();

                    foreach (var part in parts)
                    {
                        var trimmed = part.Trim();
                        if (trimmed.StartsWith("Province:"))
                            filter.Province = trimmed.Replace("Province:", "").Trim();
                        else if (trimmed.StartsWith("Technology:"))
                            filter.Technology = trimmed.Replace("Technology:", "").Trim();
                        else if (trimmed.StartsWith("Dates:"))
                            dateList.Add(trimmed.Replace("Dates:", "").Trim());
                        else if (DateTime.TryParse(trimmed, out var _))
                            dateList.Add(trimmed);
                        else if (trimmed.StartsWith("Interval:"))
                            filter.Interval = trimmed.Replace("Interval:", "").Trim();
                        else if (trimmed.StartsWith("KPI:"))
                            filter.KPI = trimmed.Replace("KPI:", "").Trim();
                        else if (trimmed.StartsWith("Threshold:"))
                            filter.Threshold = trimmed.Replace("Threshold:", "").Trim();
                        else if (trimmed.StartsWith("MinAvailability:"))
                            filter.MinAvailability = trimmed.Replace("MinAvailability:", "").Trim();
                        else if (trimmed.StartsWith("MinTraffic:"))
                            filter.MinTraffic = trimmed.Replace("MinTraffic:", "").Trim();
                    }


                    filter.Dates = dateList;
                    parsedFilters.Add(filter);
                }
                catch
                {
                    return Json(new { success = false, message = "Failed to parse filters." });
                }


            }



            // Now query the DB using parsed filters
            var results = new List<object>();
            string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                for (int f = 0; f < parsedFilters.Count; f++)
                {

                    string TBL_Part1 = "";
                    string KPI = parsedFilters[f].KPI;
                    string kpiName = "";
                    int ind = 0;
                    while (KPI[ind].ToString() != ">" && KPI[ind].ToString() != "<")
                    {
                        kpiName = kpiName + KPI[ind];
                        ind++;
                    }
                    kpiName = kpiName.Substring(0, kpiName.Length - 1);
                    string sign = KPI.Substring(ind, KPI.Length - ind);
                    string Threshold = parsedFilters[f].Threshold;
                    string Province = parsedFilters[f].Province;
                    string Technology = parsedFilters[f].Technology;
                    string Interval = parsedFilters[f].Interval;
                    string MinAva = parsedFilters[f].MinAvailability;
                    string MinTra = parsedFilters[f].MinTraffic;
                    string DateNmae = "";
                    string PIndex = "";


                    // finding expressions of CC2,CC3,RD3,RD4
                    switch (Technology)
                    {
                        case ("2G"):
                            TBL_Part1 = "CC2";
                            DateNmae = "Date";
                            break;
                        case ("3G"):
                            if (KPI.Substring(0, 2) == "CS" || KPI == "Soft_HO_SR <" || KPI == "Inter_Carrier_HO_SR <" || KPI == "Cell_Availability <")
                            {
                                TBL_Part1 = "CC3";
                                DateNmae = "Date";
                            }
                            else
                            {
                                TBL_Part1 = "RD3";
                                DateNmae = "Date";
                            }
                            break;
                        case ("4G"):
                            TBL_Part1 = "RD4";
                            DateNmae = "Datetime";
                            break;
                        default:
                            TBL_Part1 = "";
                            DateNmae = "";
                            break;
                    }


                    // Province Indexes
                    switch (Province)
                    {
                        case ("Alborz"): PIndex = "KJ"; break;
                        case ("Ardebil"): PIndex = "AR"; break;
                        case ("Bushehr"): PIndex = "BU"; break;
                        case ("Charmahal"): PIndex = "CH"; break;
                        case ("AzarSharghi"): PIndex = "AS"; break;
                        case ("Isfahan"): PIndex = "ES"; break;
                        case ("Fars"): PIndex = "FS"; break;
                        case ("Gilan"): PIndex = "GL"; break;
                        case ("Golestan"): PIndex = "GN"; break;
                        case ("Hamedan"): PIndex = "HN"; break;
                        case ("Hormozgan"): PIndex = "HZ"; break;
                        case ("Ilam"): PIndex = "IL"; break;
                        case ("Kerman"): PIndex = "KM"; break;
                        case ("Kermanshah"): PIndex = "KS"; break;
                        case ("KhorasanRazavi"): PIndex = "KH"; break;
                        case ("Khuzestan"): PIndex = "KZ"; break;
                        case ("Kohkiloyeh"): PIndex = "KB"; break;
                        case ("Kordestan"): PIndex = "KD"; break;
                        case ("Lorestan"): PIndex = "LN"; break;
                        case ("Markazi"): PIndex = "KM"; break;
                        case ("Mazandaran"): PIndex = "MA"; break;
                        case ("KhorasanShomali"): PIndex = "NK"; break;
                        case ("Qazvin"): PIndex = "QN"; break;
                        case ("Qom"): PIndex = "QM"; break;
                        case ("Semnan"): PIndex = "SM"; break;
                        case ("Sistan"): PIndex = "SB"; break;
                        case ("KhorasanJonobi"): PIndex = "SK"; break;
                        case ("Tehran"): PIndex = "TH"; break;
                        case ("AzarGharbi"): PIndex = "AG"; break;
                        case ("Yazd"): PIndex = "YZ"; break;
                        case ("Zanjan"): PIndex = "ZN"; break;
                        default: PIndex = ""; break;
                    }

                    // Table Names
                    string Ericsson_Table_Name = TBL_Part1 + "_Ericsson_Cell_" + Interval;
                    string Huawei_Table_Name = TBL_Part1 + "_Huawei_Cell_" + Interval;
                    string Nokia_Table_Name = TBL_Part1 + "_Nokia_Cell_" + Interval;


                    // Definitions of KPIs in Database (Ericsson , Huawei , Nokia) (Basis is Daily)
                    var kpiMapping2G = new Dictionary<string, List<string>>
                    {
                        { "CDR >", new List<string> { "[CDR(not Affected by incoming Handovers from 3G)(Eric_CELL)]", "[CDR3]", "[CDR(including_CS_IRAT_handovers_3G_to2G)(Nokia_SEG)]" } },
                        { "CSSR <", new List<string> { "[CSSR_MCI]", "[CSSR3]", "[CSSR_MCI]" } },
                        { "IHSR <", new List<string> { "[IHSR]", "[IHSR2]", "[IHSR]" } },
                        { "OHSR <", new List<string> { "[OHSR]", "[OHSR2]", "[OHSR]" } },
                        { "RxQual_DL <", new List<string> { "[RxQual_DL]", "[RX_QUALITTY_DL_NEW]", "[RxQuality_DL]" } },
                        { "RxQual_UL <", new List<string> { "[RxQual_UL]", "[RX_QUALITTY_UL_NEW]", "[RxQuality_UL]" } },
                        { "SDCCH_Access_SR <", new List<string> { "[SDCCH_Access_Succ_Rate]", "[SDCCH_Access_Success_Rate2]", "[SDCCH_Access_Success_Rate]" } },
                        { "SDCCH_Congestion >", new List<string> { "[SDCCH_Congestion]", "[SDCCH_Congestion_Rate]", "[SDCCH_Congestion_Rate]" } },
                        { "SDCCH_Drop_Rate >", new List<string> { "[SDCCH_Drop_Rate]", "[SDCCH_Drop_Rate]", "[SDCCH_Drop_Rate]" } },
                        { "TCH_Assign_FR >", new List<string> { "[TCH_Assign_Fail_Rate(NAK)(Eric_CELL)]", "[TCH_Assignment_FR]", "[TCH_Assignment_Failure_Rate(Nokia_SEG)]" } },
                        { "TCH_Congestion >", new List<string> { "[TCH_Congestion]", "[TCH_Cong]", "[TCH_Cong_Rate]" } },
                        { "TCH_Traffic (Erlang) <=", new List<string> { "[TCH_Traffic]", "[TCH_Traffic]", "[TCH_Traffic]" } },
                        { "TCH_Availability <", new List<string> { "[TCH_Availability]", "[TCH_Availability]", "[TCH_Availability]" } }
                    };


                    var kpiMapping3G = new Dictionary<string, List<string>>
                    {
                        { "CS_RAB_Establish <", new List<string> { "[Cs_RAB_Establish_Success_Rate]","[CS_RAB_Setup_Success_Ratio]","[CS_RAB_Establish_Success_Rate]" } },
                        { "CS_IRAT_HO_SR <", new List<string> { "[IRAT_HO_Voice_Suc_Rate]","[CS_IRAT_HO_SR]","[Inter_sys_RT_Hard_HO_SR_3Gto2G(CELL_nokia)]" } },
                        { "CS_Drop_Rate >", new List<string> { "[CS_Drop_Call_Rate]","[AMR_Call_Drop_Ratio_New(Hu_CELL)]","[CS_Drop_Call_Rate]" } },
                        { "Soft_HO_SR <", new List<string> { "[Soft_HO_Suc_Rate]","[Softer_Handover_Success_Ratio(Hu_Cell)]","[Soft_HO_Success_rate_RT]" } },
                        { "CS_RRC_SR <", new List<string> { "[CS_RRC_Setup_Success_Rate]","[CS_RRC_Connection_Establishment_SR]","[CS_RRC_SETUP_SR_WITHOUT_REPEAT(CELL_NOKIA)]" } },
                        { "CS_MultiRAB_SR <", new List<string> { "[CS_Multi_RAB_Establish_Success_Rate(Without_Nas)(CELL_Eric)]","[CSPS_RAB_Setup_Success_Ratio]","[CSAMR+PS_MRAB_STP_SR]" } },
                        { "CS_Setup_SR <", new List<string> { "[CS_Setup_Success_Rate(Without_NAS)(Cell_Eric)]","[CS_CSSR]","[Voice_Call_Setup_Success_Ratio(Nokia_CELL)]" } },
                        { "CS_RAB_Congestion_Rate >", new List<string> {"[CS_RAB_Congestion_rate(Cell_Eric)]","[CS_RAB_Setup_Congestion_Rate(Hu_Cell)]","[RAB_setup_FR_for_CS_voice_due_to_AC(CELL_nokia)]" } },
                        { "Inter_Carrier_HO_SR <", new List<string> { "[inter_frequency_handover_success_rate_for_speech(UCell_Eric)]","[InterFrequency_Hardhandover_success_Ratio_CSservice]","[Intra_RNC_Inter_frequency_HO_Success_Rate_RT]" } },
                        { "CS_Traffic (Erlang) <=", new List<string> { "[CS_Traffic]","[CS_Erlang]","[CS_Traffic]" } },
                        { "Cell_Availability <", new List<string> { "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]","[Radio_Network_Availability_Ratio(Hu_Cell)]","[Cell_Availability_excluding_blocked_by_user_state]" } },
                        { "HSDPA_SR <", new List<string> {"[HSDPA_RAB_Setup_Succ_Rate(UCell_Eric)]","[HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)]","[HSDPA_setup_success_ratio_from_user_perspective(CELL_Nokia)]" } },
                        { "HSUPA_SR <", new List<string> {"[HSUPA_Setup_Success_Rate(UCell_Eric)]","[HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)]","[HSUPA_Setup_Success_Ratio_from_user_perspective(CELL)]" } },
                        { "DL_User_THR (Mbps) <", new List<string> { "[HS_USER_Throughput_NET_PQ(Mbps)(UCell_Eric)]","[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)]","[AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(Nokia_CELL)]" } },
                        { "HSDAP_Drop_Rate >", new List<string> { "[HSDPA_Drop_Call_Rate(UCell_Eric)]","[HSDPA_cdr(%)_(Hu_Cell)_new]","[HSDPA_Call_Drop_Rate(Nokia_Cell)]" } },
                        { "HSUAP_Drop_Rate >", new List<string> { "[HSUPA_Drop_Call_Rate(UCell_Eric)]","[HSUPA_CDR(%)_(Hu_Cell)_new]","[HSUPA_Call_Drop_Rate(Nokia_CELL)]" } },
                        { "PS_RRC_SR <", new List<string> {"[PS_RRC_Setup_Success_Rate(UCell_Eric)]","[PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)]","[PS_RRCSETUP_SR]" } },
                        { "Ps_RAB_Establish <", new List<string> {"[Ps_RAB_Establish_Success_Rate(UCell_Eric)]","[PS_RAB_Setup_Success_Ratio(Hu_Cell)]","[RAB_Setup_and_Access_Complete_Ratio_for_NRT_Service_from_User_pe]" } },
                        { "PS_MultiRAB_Establish <", new List<string> {"[PS_Multi_RAB_Establish_Success_Rate(without_Nas)(UCELL_Eric)]","[CS+PS_RAB_Setup_Success_Ratio]","[CSAMR+PS_MRAB_stp_SR(Nokia_CELL)]" } },
                        { "PS_Drop_Rate >", new List<string> {"[PS_Drop_Call_Rate(UCell_Eric)]","[PS_Call_Drop_Ratio]","[Packet_Session_Drop_Ratio_NOKIA(CELL_NOKIA)]"} },
                        { "HSDPA_Cell_Change_SR <", new List<string> { "[HSDPA_Cell_Change_Succ_Rate(UCell_Eric)]","[HSDPA_Soft_HandOver_Success_Ratio]","[HSDPA_Cell_Change_SR(Nokia_CELL)]" } },
                        { "HS_Share_Payload <", new List<string> {"[HS_share_PAYLOAD_Rate(UCell_Eric)]","[HS_share_PAYLOAD_%]","[HS_SHARE_PAYLOAD(Nokia_CELL)]"} },
                        { "DL_Cell_THR (Mbps) <", new List<string> {"[HSDPA_Cell_Scheduled_Throughput(mbps)(UCell_Eric)]","[HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)]","[Active_HS-DSCH_cell_throughput_mbs(CELL_nokia)]" } },
                        { "RSSI (dBm) >", new List<string> {"[uplink_average_RSSI_dbm_(Eric_UCELL)]","[Mean_RTWP(Cell_Hu)]","[average_RTWP_dbm(Nokia_Cell)]" } },
                        { "Average CQI <", new List<string> {"[Avg_CQI(UCell_Eric)]","[CQI_new(Hu_Cell)]","[AVERAGE_CQI(cell_nokia)]" } },
                        { "PS_Payload (GB) <=", new List<string> { "[PS_Volume(GB)(UCell_Eric)]","[PAYLOAD]","[PS_Payload_Total(HS+R99)(Nokia_CELL)_GB]" } }
                    };


                    var kpiMapping4G = new Dictionary<string, List<string>>
                    {
                        { "RRC_Connection_SR <", new List<string> {"[RRC_Estab_Success_Rate(ReAtt)(EUCell_Eric)]","[RRC_Connection_Setup_Success_Rate_service]","[RRC_Connection_Setup_Success_Ratio(Nokia_LTE_CELL)]" } },
                        { "ERAB_SR_Initial <", new List<string> {"[Initial_ERAB_Estab_Success_Rate(eNodeB_Eric)]","[E-RAB_Setup_Success_Rate]","[Initial_E-RAB_Setup_Success_Ratio(Nokia_LTE_CELL)]" } },
                        { "ERAB_SR_Added <", new List<string> { "[E-RAB_Setup_SR_incl_added_New(EUCell_Eric)]","[E-RAB_Setup_Success_Rate(Hu_Cell)]","[E-RAB_Setup_SR_incl_added(Nokia_LTE_CELL)]" } },
                        { "DL_THR (Mbps) <", new List<string> {"[Average_UE_DL_Throughput(Mbps)(eNodeB_Eric)]","[Average_Downlink_User_Throughput(Mbit/s)]","[User_Throughput_DL_mbps(Nokia_LTE_CELL)]" } },
                        { "UL_THR (Mbps) <", new List<string> {"[Average_UE_UL_Throughput(Mbps)(eNodeB_Eric)]","[Average_UPlink_User_Throughput(Mbit/s)]","[User_Throughput_UL_mbps(Nokia_LTE_CELL)]"} },
                        { "ERAB_Drop_Rate >", new List<string> {"[E_RAB_Drop_Rate(eNodeB_Eric)]","[Call_Drop_Rate]","[E-RAB_Drop_Ratio_RAN_View(Nokia_LTE_CELL)_NEW]"} },
                        { "S1_Signalling_SR <", new List<string> {"[S1Signal_Estab_Success_Rate(EUCell_Eric)]","[S1Signal_E-RAB_Setup_SR(Hu_Cell)]","[S1Signal_E-RAB_Setup_SR(Nokia_LTE_CELL)]" } },
                        { "Intra_Freq_SR <", new List<string> { "[IntraF_Handover_Execution(eNodeB_Eric)]","[IntraF_HOOut_SR]","[HO_Success_Ratio_intra_eNB(Nokia_LTE_CELL)]" } },
                        { "Inter_Freq_SR <", new List<string> {"[InterF_Handover_Execution(eNodeB_Eric)]","[InterF_HOOut_SR]","[Inter-Freq_HO_SR(Nokia_LTE_CELL)]" } },
                        { "UL_Packet_Loss >", new List<string> {"[Average_UE_Ul_Packet_Loss_Rate(eNodeB_Eric)]","[Average_UL_Packet_Loss_%(Huawei_LTE_UCell)]","[Packet_loss_UL(Nokia_EUCELL)]" } },
                        { "UE_DL_Latency (ms) >", new List<string> {"[Average_UE_DL_Latency(ms)(eNodeB_Eric)]","[Average_DL_Latency_ms(Huawei_LTE_EUCell)]","[Average_Latency_DL_ms(Nokia_LTE_CELL)]" } },
                        { "Average_CQI <", new List<string> {"[CQI_(EUCell_Eric)]","[Average_CQI(Huawei_LTE_Cell)]","[Average_CQI(Nokia_LTE_CELL)]" } },
                        { "PUCCH_RSSI (dBm) >", new List<string> {"[RSSI_PUCCH(EUCell_Eric)]","[RSSI_PUCCH(Huawei_LTE_Cell)]","[Average_RSSI_for_PUCCH(Nokia_LTE_CELL)]" } },
                        { "PUSCH_RSSI (dBm) >", new List<string> {"[RSSI_PUSCH(EUCell_Eric)]","[RSSI_PUSCH(Huawei_LTE_Cell)]","[Average_RSSI_for_PUSCH(Nokia_LTE_CELL)]" } },
                        { "Total_Paylaod (GB) <=", new List<string> {"[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]","[Total_Traffic_Volume(GB)]","[Total_Payload_GB(Nokia_LTE_CELL)]" } },
                        { "Cell_Availability <", new List<string> { "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]", "[Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]", "[cell_availability_include_manual_blocking(Nokia_LTE_CELL)]" } }
                    };


                    // Selected KPIs
                    var selectedKPIs = new List<string>();

                    if (kpiMapping2G.ContainsKey(KPI))
                    {
                        selectedKPIs = kpiMapping2G[KPI];
                    }
                    else if (kpiMapping3G.ContainsKey(KPI))
                    {
                        selectedKPIs = kpiMapping3G[KPI];
                    }
                    else if (kpiMapping4G.ContainsKey(KPI))
                    {
                        selectedKPIs = kpiMapping4G[KPI];
                    }

                    // Query
                    if (MinTra == "")
                    {
                        MinTra = "0";
                    }
                    if (MinAva == "")
                    {
                        MinAva = "0";
                    }

                    string StringDate = "";
                    for (int t = 0; t < parsedFilters[f].Dates.Count; t++)
                    {
                        StringDate = StringDate + "cast(" + DateNmae + " as Date)='" + parsedFilters[f].Dates[t].ToString() + "' or ";
                    }
                    StringDate = StringDate.Substring(0, StringDate.Length - 4);

                    string StringDate2 = "";
                    for (int t = 0; t < parsedFilters[f].Dates.Count; t++)
                    {
                        StringDate2 = StringDate2 + "cast(Date as Date)='" + parsedFilters[f].Dates[t].ToString() + "' or ";
                    }
                    StringDate2 = StringDate2.Substring(0, StringDate2.Length - 4);



                    if (Technology == "2G")
                    {
                        string query = "";
                        if (Interval == "Daily")
                        {
                            if (selectedKPIs[0] == "[TCH_Traffic]" || selectedKPIs[0] == "[TCH_Availability]")
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Ericsson_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Huawei_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Seg  as 'Cell', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Nokia_Table_Name + " where substring(Seg,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else // Query for all KPIs except traffic and availability
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Ericsson_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Huawei_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Seg  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', TCH_Traffic as 'Traffic', TCH_Availability as 'Availability' from " + Nokia_Table_Name + " where substring(Seg,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and TCH_Traffic>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ")";
                            }

                        }
                        else
                        {
                            if (selectedKPIs[0] == "[TCH_Traffic]" || selectedKPIs[0] == "[TCH_Availability]")
                            {
                                if (selectedKPIs[0] == "[TCH_Traffic]")    // BH Traffic KPI names are usually different and need to be changed
                                {
                                    selectedKPIs[0] = "[TCH_Traffic_BH]";
                                    selectedKPIs[1] = "[TCH_Traffic_BH]";
                                    selectedKPIs[2] = "[TCH_Traffic_BH]";
                                }
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Ericsson_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Huawei_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Seg  as 'Cell', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Nokia_Table_Name + " where substring(Seg,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else  // Query for all KPIs except traffic and availability
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Ericsson_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Cell as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Huawei_Table_Name + " where substring(Cell,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'2G' as 'Technology', BSC as 'Node', '" + Interval + "' as 'Interval', Seg  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', TCH_Traffic_BH as 'Traffic', TCH_Availability as 'Availability' from " + Nokia_Table_Name + " where substring(Seg,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and TCH_Traffic_BH>=" + MinTra + " and TCH_Availability>=" + MinAva + " and (" + StringDate + ")";
                            }

                        }

                        Data_Table_2G = new DataTable();
                        Data_Table_2G = Query_Execution_Table_Output(query);

                    }
                    if (Technology == "3G")
                    {
                        string query = "";
                        if (Interval == "Daily" && TBL_Part1 == "CC3")
                        {
                            if (selectedKPIs[0] == "[CS_Traffic]" || selectedKPIs[0] == "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]")
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', CS_Traffic as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and CS_Traffic>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', CS_Erlang as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and CS_Erlang>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell', CS_Traffic as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and CS_Traffic>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state]>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else 
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', CS_Traffic as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and CS_Traffic>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', CS_Erlang as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and CS_Erlang>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', CS_Traffic as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and CS_Traffic>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state]>=" + MinAva + " and (" + StringDate + ")";
                            }

                        }
                        if (Interval == "BH" && TBL_Part1 == "CC3")
                        {
                            if (selectedKPIs[0] == "[CS_Traffic_BH]" || selectedKPIs[0] == "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]")
                            {
                                if (selectedKPIs[0] == "[CS_Traffic_BH]")
                                {
                                    selectedKPIs[0] = "[TCH_Traffic_BH]";
                                    selectedKPIs[1] = "[CS_Erlang]";
                                    selectedKPIs[2] = "[CS_TrafficBH]";
                                }
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', CS_Traffic_BH as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and CS_Traffic_BH>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', CS_Erlang as 'Traffic',  [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and CS_Erlang>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell', CS_TrafficBH as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and CS_TrafficBH>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state]>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', CS_Traffic_BH as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and CS_Traffic_BH>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', CS_Erlang as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and CS_Erlang>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', CS_TrafficBH as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and CS_TrafficBH>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state]>=" + MinAva + " and (" + StringDate + ")";
                            }
                        }
                        if (Interval == "Daily" && TBL_Part1 == "RD3")
                        {
                            if (selectedKPIs[0] == "[PS_Volume(GB)(UCell_Eric)]" || selectedKPIs[0] == "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]")
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', [PS_Volume(GB)(UCell_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [PS_Volume(GB)(UCell_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', PAYLOAD as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and PAYLOAD>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB]>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', [PS_Volume(GB)(UCell_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [PS_Volume(GB)(UCell_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', PAYLOAD as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and PAYLOAD>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB] as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [PS_Payload_Total(HS+R99)(Nokia_CELL)_GB]>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]>=" + MinAva + " and (" + StringDate + ")";
                            }


                        }
                        if (Interval == "BH" && TBL_Part1 == "RD3")
                        {
                            if (selectedKPIs[0] == "[PS_Volume(GB)(UCell_Eric)]" || selectedKPIs[0] == "[Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]")
                            {
                                if (selectedKPIs[0] == "[PS_Volume(GB)(UCell_Eric)]")
                                {
                                    selectedKPIs[0] = "[Payload_Total_BH]";
                                    selectedKPIs[1] = "[Payload_Total_BH]";
                                    selectedKPIs[2] = "[Payload_Total_BH]";
                                }
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', Payload_Total_BH as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell', Payload_Total_BH as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell', Payload_Total_BH as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else
                            {
                                query = @"SELECT Date, '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', Payload_Total_BH as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)] as 'Availability' from " + Ericsson_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(UCELL_Eric)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1 as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', Payload_Total_BH as 'Traffic', [Radio_Network_Availability_Ratio(Hu_Cell)] as 'Availability' from " + Huawei_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Radio_Network_Availability_Ratio(Hu_Cell)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                        "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'3G' as 'Technology', ElementID as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', Payload_Total_BH as 'Traffic', [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and Payload_Total_BH>=" + MinTra + " and [Cell_Availability_excluding_blocked_by_user_state(Nokia_UCell)]>=" + MinAva + " and (" + StringDate + ")";
                            }
                        }
                        Data_Table_3G = new DataTable();
                        Data_Table_3G = Query_Execution_Table_Output(query);

                    }
                    if (Technology == "4G")
                    {
                        string query = "";
                        if (Interval == "Daily")
                        {
                        
                            if (selectedKPIs[0] == "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]" || selectedKPIs[0] == "[Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]")
                            {
                                query = @"SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Availability' from " + Ericsson_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [Total_Volume(UL+DL)(GB)(eNodeB_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell', [Total_Traffic_Volume(GB)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Availability' from " + Huawei_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and [Total_Traffic_Volume(GB)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell', [Total_Payload_GB(Nokia_LTE_CELL)] as 'Traffic', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [Total_Payload_GB(Nokia_LTE_CELL)]>=" + MinTra + " and [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]>=" + MinAva + " and (" + StringDate2 + ")";
                            }
                            else // Query for all KPIs except traffic and availability
                            {
                                query = @"SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Availability' from " + Ericsson_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [Total_Volume(UL+DL)(GB)(eNodeB_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', [Total_Traffic_Volume(GB)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Availability' from " + Huawei_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and [Total_Traffic_Volume(GB)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Date, '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', ElementID1  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', [Total_Payload_GB(Nokia_LTE_CELL)] as 'Traffic', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)] as 'Availability' from " + Nokia_Table_Name + " where substring(ElementID1,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [Total_Payload_GB(Nokia_LTE_CELL)]>=" + MinTra + " and [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]>=" + MinAva + " and (" + StringDate2 + ")";
                            }
                        }
                        else
                        {
                            if (selectedKPIs[0] == "[Total_Volume(UL+DL)(GB)(eNodeB_Eric)]" || selectedKPIs[0] == "[Cell_Availability_Rate_Include_Blocking(Cell_EricLTE)]")
                            {
                                query = @"SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Availability' from " + Ericsson_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [Total_Volume(UL+DL)(GB)(eNodeB_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell', [Total_Traffic_Volume(GB)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Availability' from " + Huawei_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and [Total_Traffic_Volume(GB)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB  as 'Cell', [Total_Payload_GB(Nokia_LTE_CELL)] as 'Traffic', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)] as 'Availability' from " + Nokia_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [Total_Payload_GB(Nokia_LTE_CELL)]>=" + MinTra + " and [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]>=" + MinAva + " and (" + StringDate + ")";
                            }
                            else // Query for all KPIs except traffic and availability
                            {
                                query = @"SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Ericsson' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[0] + " as 'KPI Value', [Total_Volume(UL+DL)(GB)(eNodeB_Eric)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)] as 'Availability' from " + Ericsson_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[0] + sign + Threshold + " and [Total_Volume(UL+DL)(GB)(eNodeB_Eric)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_EricLTE)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Huawei' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[1] + " as 'KPI Value', [Total_Traffic_Volume(GB)] as 'Traffic', [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)] as 'Availability' from " + Huawei_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[1] + sign + Threshold + " and [Total_Traffic_Volume(GB)]>=" + MinTra + " and [Cell_Availability_Rate_Exclude_Blocking(Cell_Hu)]>=" + MinAva + " and (" + StringDate + ") union all " +
                                          "SELECT Datetime as 'Date', '" + Province + "' as 'Province', 'Nokia' as 'Vendor'" + " ,'4G' as 'Technology', '' as 'Node', '" + Interval + "' as 'Interval', eNodeB  as 'Cell'" + ", '" + kpiName + "' as 'KPI Name', " + selectedKPIs[2] + " as 'KPI Value', [Total_Payload_GB(Nokia_LTE_CELL)] as 'Traffic', [cell_availability_include_manual_blocking(Nokia_LTE_CELL)] as 'Availability' from " + Nokia_Table_Name + " where substring(eNodeB,1,2)='" + PIndex + "' and " + selectedKPIs[2] + sign + Threshold + " and [Total_Payload_GB(Nokia_LTE_CELL)]>=" + MinTra + " and [cell_availability_include_manual_blocking(Nokia_LTE_CELL)]>=" + MinAva + " and (" + StringDate + ")";
                            }

                        }

                        Data_Table_4G = new DataTable();
                        Data_Table_4G = Query_Execution_Table_Output(query);
                    }

                    mergedTable2G.Merge(Data_Table_2G);
                    mergedTable3G.Merge(Data_Table_3G);
                    mergedTable4G.Merge(Data_Table_4G);
                }

                mergedTable.Merge(mergedTable2G);
                mergedTable.Merge(mergedTable3G);
                mergedTable.Merge(mergedTable4G);
            }

            var dataList = mergedTable.AsEnumerable()
    .Select(row => mergedTable.Columns.Cast<DataColumn>()
        .ToDictionary(col => col.ColumnName, col => row[col]?.ToString()))
    .ToList();

            return Json(new { success = true, data = dataList });
        }

    }
}
