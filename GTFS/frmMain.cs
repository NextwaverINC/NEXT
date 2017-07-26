using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using interop = System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace GTFS
{
    public partial class frmMain : Form
    {
        #region Excel
        Excel.Application excelApp;
        object missing = Type.Missing;
        Excel.Sheets excelSheets;
        Excel.Workbook excelWorkbook;
        Excel.Worksheet SheetData;
        public bool OpenExcel(string ExcelFileName)
        {
            try
            {
                excelApp = new Excel.Application();
                excelApp.EnableEvents = false;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(1033);
                // Creates a new Excel Application

                // Makes Excel visible to the user.
                // excelApp.Visible = true;
                Excel.Workbook newWorkbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                // The following code opens an existing workbook            

                excelWorkbook = excelApp.Workbooks.Open(ExcelFileName, 0,
                    false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,
                    false, 0, true, false, false);

                // The following gets the Worksheets collection
                excelSheets = excelWorkbook.Worksheets;
                SheetData = (Excel.Worksheet)excelSheets[1];
                String Namesss = SheetData.Name;
                return true;
            }
            catch (Exception excelApp) { 
                CloseExcel();
                alError.Add("OpenExcel => FilePath:" + ExcelFileName + ", Error:" + excelApp.Message);
                return false; 
            }
        }
        public void SaveExcel(string PathSaveFile)
        {
            
            excelWorkbook.SaveAs(PathSaveFile, Excel.XlFileFormat.xlWorkbookNormal, null,
                            null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null,
                            null);
          
        }
        public void CloseExcel()
        {
            if (excelApp != null)
            {
                //excelWorkbook.Saved = false;
                //excelWorkbook.RunAutoMacros(XlRunAutoMacro.xlAutoClose);

                //excelApp.DisplayAlerts = false;
                //excelApp.DisplayInfoWindow = false;
                //excelApp.DisplayDocumentActionTaskPane = false;
                excelApp.EnableEvents = false;
                //excelWorkbook.Save();
                if (excelWorkbook != null)
                    excelWorkbook.Close(false, missing, false);
                excelApp.EnableEvents = true;
                excelApp.Quit();
                excelApp = null;
            }            
        }

        bool CheckAgency(string str)
        {
            switch (str.Trim())
            {
                case "A001": return true;
                case "A002": return true;
                case "A003": return true;
                case "A004": return true;
                case "A005": return true;
                case "A006": return true;
                case "A007": return true;
                case "A008": return true;
                case "A009": return true;
            }

            return false;
        }

        /// <summary>
        /// line,stop,stop_times,prices
        /// </summary>
        /// <param name="PathFile"></param>
        /// <returns></returns>
        public bool LoadTemplate(string PathFile, bool is_inbound)
        {
            if (OpenExcel(PathFile))
            {
                try
                {
                    string Agency = "" + SheetData.get_Range("B2", missing).Value2;
                    string Sql;
                    string ErrorMSG;
                    if (CheckAgency(Agency))
                    {
                        string Line = "" + SheetData.get_Range("B3", missing).Value2;
                        string Way = "" + SheetData.get_Range("B4", missing).Value2;

                        string Work_Day = ("" + SheetData.get_Range("B5", missing).Value2).Trim();
                        Work_Day = Work_Day.Replace(" ", "");
                        Sql = @"SELECT [service_id]
                                FROM [Google_Transit].[dbo].[calendar]
                                WHERE [code] = '" + Work_Day.Trim() + "'";
                        System.Data.DataTable dtTmp = Retreive(Sql);
                        string service_id = "SV0001";
                        if (dtTmp.Rows.Count != 0)
                            service_id = "" + dtTmp.Rows[0]["service_id"];

                        int Round_Count = strToInt("" + SheetData.get_Range("E5", missing).Value2);
                        int Station_Count = strToInt("" + SheetData.get_Range("G5", missing).Value2);
                        string Work_Time = ("" + SheetData.get_Range("B6", missing).Value2).Trim();
                        string Detail = ("" + SheetData.get_Range("B7", missing).Value2).Trim();

                        Sql = @"SELECT [line_id]
                          FROM [Google_Transit].[dbo].[line]
                          WHERE [line_code] = '" + Line + @"'
                          AND [agency_id] = '" + Agency + @"'";
                        dtTmp = Retreive(Sql);
                        string tmp_line_id = "";
                        if (dtTmp.Rows.Count != 0)
                        {
                            tmp_line_id = "" + dtTmp.Rows[0]["line_id"];
                            if (is_inbound)
                            {
                                Sql = @"UPDATE [Google_Transit].[dbo].[line]
                                    SET work_day = '" + Work_Day + @"'
                                       ,work_time = '" + Work_Time + @"'
                                       ,detail = '" + Detail + @"'
                                       ,station_in_count = " + Station_Count + @"
                                       ,round_count = " + Round_Count + @"
                                       ,file_detail = '" + PathFile + @"'
                                       ,inbound_detail = '" + Way + @"'
                                      WHERE line_id = '" + tmp_line_id + "'";
                            }
                            else
                            {
                                Sql = @"UPDATE [Google_Transit].[dbo].[line]
                                    SET work_day = '" + Work_Day + @"'
                                       ,work_time = '" + Work_Time + @"'
                                       ,detail = '" + Detail + @"'
                                       ,station_out_count = " + Station_Count + @"
                                       ,round_count = " + Round_Count + @"
                                       ,file_detail = '" + PathFile + @"'
                                       ,outbound_detail = '" + Way + @"'
                                      WHERE line_id = '" + tmp_line_id + "'";
                            }
                            Execute(Sql, out ErrorMSG);
                        }
                        else
                        {
                            Sql = @"SELECT CONVERT(decimal(18,0),ISNULL(SUBSTRING(MAX([line_id]),7,LEN(MAX([line_id]))),0)) + 1 as new_id
                                      FROM [Google_Transit].[dbo].[line]
                                      WHERE SUBSTRING([line_id],1,6) = 'LI" + Agency + "'";

                            dtTmp = Retreive(Sql);
                            tmp_line_id = "LI" + Agency + genId("" + dtTmp.Rows[0]["new_id"], 6);

                            if (is_inbound)
                            {
                                Sql = @"INSERT INTO [Google_Transit].[dbo].[line]
                                               ([line_id]
                                               ,[line_code]
                                               ,[agency_id]
                                               ,[inbound_detail]
                                               ,[work_day]
                                               ,[work_time]
                                               ,[detail]                              
                                               ,[file_detail]
                                               ,[station_in_count]
                                               ,[round_count])
                                         VALUES
                                               ('@line_id'
                                               ,'@line_code'
                                               ,'@agency_id'  
                                               ,'@inbound_detail'
                                               ,'@work_day'
                                               ,'@work_time'
                                               ,'@detail'
                                               ,'@file_detail'
                                               ,'@station_count'
                                               ,'@round_count')";
                            }
                            else
                            {
                                Sql = @"INSERT INTO [Google_Transit].[dbo].[line]
                                           ([line_id]
                                           ,[line_code]
                                           ,[agency_id]
                                           ,[outbound_detail]   
                                           ,[work_day]
                                           ,[work_time]
                                           ,[detail]                              
                                           ,[file_detail]
                                           ,[station_out_count]
                                           ,[round_count])
                                     VALUES
                                           ('@line_id'
                                           ,'@line_code'
                                           ,'@agency_id'  
                                           ,'@outbound_detail'
                                           ,'@work_day'
                                           ,'@work_time'
                                           ,'@detail'
                                           ,'@file_detail'
                                           ,'@station_count'
                                           ,'@round_count')";
                            }


                            Sql = Sql.Replace("@line_id", tmp_line_id);
                            Sql = Sql.Replace("@line_code", Line);
                            Sql = Sql.Replace("@agency_id", Agency);
                            Sql = Sql.Replace("@inbound_detail", Way);
                            Sql = Sql.Replace("@outbound_detail", Way);
                            Sql = Sql.Replace("@file_detail", PathFile);
                            Sql = Sql.Replace("@work_day", Work_Day);
                            Sql = Sql.Replace("@work_time", Work_Time);
                            Sql = Sql.Replace("@detail", Detail);
                            Sql = Sql.Replace("@station_count", "" + Station_Count);
                            Sql = Sql.Replace("@round_count", "" + Round_Count);

                            if (!Execute(Sql, out ErrorMSG))
                            {
                                alError.Add("INSERT line => line_id:" + tmp_line_id + ",line_code:" + Line + ",Error:" + ErrorMSG);
                                CloseExcel();
                                return false;
                            }
                        }

                        int S_INDEX = 12;
                        string[] listColumn = {"I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z" 
                       ,"AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"
                       ,"BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ"
                       ,"CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"
                       ,"DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DX","DY","DZ"
                       ,"EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM","EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ"
                       ,"FA","FB","FC","FD","FE","FF","FG","FH","FI","FJ","FK","FL","FM","FN","FO","FP","FQ","FR","FS","FT","FU","FV","FW","FX","FY","FZ"
                       ,"GA","GB","GC","GD","GE","GF","GG","GH","GI","GJ","GK","GL","GM","GN","GO","GP","GQ","GR","GS","GT","GU","GV","GW","GX","GY","GZ"
                       ,"HA","HB","HC","HD","HE","HF","HG","HH","HI","HJ","HK","HL","HM","HN","HO","HP","HQ","HR","HS","HT","HU","HV","HW","HX","HY","HZ"
                       ,"IA","IB","IC","ID","IE","IF","IG","IH","II","IJ","IK","IL","IM","IN","IO","IP","IQ","IR","IS","IT","IU","IV","IW","IX","IY","IZ"};

                        System.Data.DataTable dtStop = new System.Data.DataTable();
                        dtStop.Columns.Add("STA_NO");
                        dtStop.Columns.Add("STOP_ID");

                        System.Data.DataTable dtRound = new System.Data.DataTable();
                        dtRound.Columns.Add("STA_NO");
                        dtRound.Columns.Add("ROUND");
                        dtRound.Columns.Add("TO");
                        dtRound.Columns.Add("FROM");

                        System.Data.DataTable dtPrice = new System.Data.DataTable();
                        dtPrice.Columns.Add("STA_NO1");
                        dtPrice.Columns.Add("STA_NO2");
                        dtPrice.Columns.Add("PRICE");
                        for (int i = 0; i < Station_Count; i++)
                        {
                            string STA_NO = "" + SheetData.get_Range("A" + (i + S_INDEX), missing).Value2;
                            string STA_NM_TH = "" + SheetData.get_Range("B" + (i + S_INDEX), missing).Value2;

                            STA_NM_TH = STA_NM_TH.Trim();

                            if (STA_NM_TH == "")
                                alError.Add("STOP NAME IS NULL => PathFile:" + PathFile);

                            if (Agency == "A009" && !STA_NM_TH.Contains("สถานีรถไฟ"))
                            {
                                STA_NM_TH = "สถานีรถไฟ " + STA_NM_TH;
                            }

                            string STA_NM_EN = "" + SheetData.get_Range("C" + (i + S_INDEX), missing).Value2;
                            string DISTANCE = "" + SheetData.get_Range("D" + (i + S_INDEX), missing).Value2;
                            string VELOCITY = "" + SheetData.get_Range("E" + (i + S_INDEX), missing).Value2;
                            string TIME_USE = "" + SheetData.get_Range("F" + (i + S_INDEX), missing).Value2;
                            string TIME_FM = "" + SheetData.get_Range("G" + (i + S_INDEX), missing).Value2;
                            string TEMP_PARK = "" + SheetData.get_Range("H" + (i + S_INDEX), missing).Value2;

                            STA_NM_TH = STA_NM_TH.Replace("[ INBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ OUTBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[INBOUND]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[OUTBOUND]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ INBOUND  ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ OUTBOUND  ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[  INBOUND  ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[  OUTBOUND  ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[  INBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[  OUTBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[OUTBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("INBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ INBOUND]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ OUTBOUND ]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ OUTBOUND]", "");
                            STA_NM_TH = STA_NM_TH.Replace("[ OUTBOUND ]", "");
                           
                           

                            STA_NM_TH = STA_NM_TH.Trim();

                            Sql = @"  SELECT [stop_id]
                                          FROM [Google_Transit].[dbo].[stop]
                                          WHERE replace([stop_name_th],' ','') = '" + STA_NM_TH.Replace(" ", "") + @"'";

                            dtTmp = Retreive(Sql);

                            string STOP_ID = "ST" + Agency + "_" + STA_NO;
                            if (dtTmp.Rows.Count != 0)
                            {
                                STOP_ID = "" + dtTmp.Rows[0]["stop_id"];
                                Sql = @"UPDATE [Google_Transit].[dbo].[stop]
                                    SET stop_distance = '" + DISTANCE + @"'
                                       ,stop_velocity = '" + VELOCITY + @"'
                                       ,time_use = '" + TIME_USE + @"'
                                       ,time_use_format = '" + TIME_FM + @"'
                                       ,temp_park = '" + TEMP_PARK + @"'
                                       ,stop_desc = 'name:" + PathFile + @"'
                                       ,stop_name_en = '" + STA_NM_EN + @"'
                                      WHERE stop_id = '" + STOP_ID + "'";
                                Execute(Sql, out ErrorMSG);
                            }
                            else
                            {
                                Sql = @"SELECT CONVERT(decimal(18,0),ISNULL(SUBSTRING(MAX([stop_id]),7,LEN(MAX([stop_id]))),0)) + 1 as new_id
                                      FROM [Google_Transit].[dbo].[stop]
                                      WHERE SUBSTRING([stop_id],1,6) = 'ST" + Agency + "'";

                                dtTmp = Retreive(Sql);
                                STOP_ID = "ST" + Agency + genId("" + dtTmp.Rows[0]["new_id"], 6);
                                Sql = @"INSERT INTO [Google_Transit].[dbo].[stop]
                                           ([stop_id]             
                                           ,[stop_name_th]    
                                           ,[stop_name_en]    
                                           ,[stop_desc] 
                                           ,[stop_type]
                                           ,[stop_distance]
                                           ,[stop_velocity]
                                           ,[time_use]
                                           ,[time_use_format]
                                           ,[temp_park]
                                           ,[line_code])
                                     VALUES
                                           ('@stop_id'    
                                           ,'@stop_name_th'   
                                           ,'@stop_name_en'       
                                           ,'@stop_desc'     
                                           ,'@stop_distance'
                                           ,'@stop_velocity'
                                           ,'@time_use'
                                           ,'@time_use_format'
                                           ,'@temp_park' 
                                           ,'@stop_type'
                                           ,'@line_code')";
                                Sql = Sql.Replace("@stop_id", STOP_ID);
                                Sql = Sql.Replace("@stop_name_th", STA_NM_TH);
                                Sql = Sql.Replace("@stop_name_en", STA_NM_EN);
                                Sql = Sql.Replace("@stop_desc", "name:" + PathFile);
                                Sql = Sql.Replace("@stop_distance", DISTANCE);
                                Sql = Sql.Replace("@stop_velocity", VELOCITY);
                                Sql = Sql.Replace("@time_use_format", TIME_FM);
                                Sql = Sql.Replace("@time_use", TIME_USE);
                                Sql = Sql.Replace("@temp_park", TEMP_PARK);
                                Sql = Sql.Replace("@stop_type", Agency);
                                Sql = Sql.Replace("@line_code", tmp_line_id);
                                if (!Execute(Sql, out ErrorMSG))
                                    alError.Add("INSERT stop => stop_id:" + STOP_ID + ",stop_name_th:" + STA_NM_TH + ",Error:" + ErrorMSG);
                            }

                            DataRow DR = dtStop.NewRow();
                            DR.BeginEdit();
                            DR["STA_NO"] = STA_NO;
                            DR["STOP_ID"] = STOP_ID;
                            DR.EndEdit();
                            dtStop.Rows.Add(DR);

                            int tmpIndex = 0;

                            for (int j = 0; j < Round_Count; j++)
                            {
                                string TO = "" + SheetData.get_Range(listColumn[tmpIndex] + (i + S_INDEX), missing).Text;
                                tmpIndex++;
                                string FROM = "" + SheetData.get_Range(listColumn[tmpIndex] + (i + S_INDEX), missing).Text;
                                tmpIndex++;

                                DR = dtRound.NewRow();
                                DR.BeginEdit();
                                DR["STA_NO"] = STA_NO;
                                DR["ROUND"] = "" + (j + 1);
                                DR["TO"] = TO;
                                DR["FROM"] = FROM;
                                DR.EndEdit();
                                dtRound.Rows.Add(DR);
                            }

                            if (cbxPrice.Checked)
                            {
                                for (int j = 0; j < Station_Count; j++)
                                {
                                    string PRICE = ("" + SheetData.get_Range(listColumn[tmpIndex] + (i + S_INDEX), missing).Value2).Trim();
                                    tmpIndex++;

                                    int chSTA_NO = strToInt(STA_NO);

                                    string TMP_N01 = "" + (j + 1);

                                    if (PRICE != "" && PRICE != "0" && TMP_N01 != STA_NO)
                                    {
                                        DR = dtPrice.NewRow();
                                        DR.BeginEdit();
                                        DR["STA_NO1"] = TMP_N01;
                                        DR["STA_NO2"] = STA_NO;
                                        DR["PRICE"] = PRICE;
                                        DR.EndEdit();
                                        dtPrice.Rows.Add(DR);
                                    }
                                }
                            }
                        }

                        CloseExcel();
                        //================================================================

                        for (int i = 0; i < dtRound.Rows.Count; i++)
                        {
                            string STA_NO = "" + dtRound.Rows[i]["STA_NO"];
                            string ROUND = "" + dtRound.Rows[i]["ROUND"];
                            string TO = "" + dtRound.Rows[i]["TO"];
                            string FROM = "" + dtRound.Rows[i]["FROM"];

                            TO = TO.Replace(".", ":");
                            FROM = FROM.Replace(".", ":");

                            if (TO.Trim() == "0")
                                TO = "-";
                            if (FROM.Trim() == "0")
                                FROM = "-";

                            DataRow[] DR = dtStop.Select("STA_NO = '" + STA_NO + "'");
                            string STOP_ID = "" + DR[0]["STOP_ID"];
                            if (is_inbound)
                                Sql = @"DELETE [Google_Transit].[dbo].[stop_times] WHERE [line_id] = '" + tmp_line_id + "' AND [round_no] = '" + ROUND + "' AND [stop_id] = '" + STOP_ID + "' AND is_inbound = 1 AND service_id = '" + service_id + "'";
                            else
                                Sql = @"DELETE [Google_Transit].[dbo].[stop_times] WHERE [line_id] = '" + tmp_line_id + "' AND [round_no] = '" + ROUND + "' AND [stop_id] = '" + STOP_ID + "' AND is_inbound = 0 AND service_id = '" + service_id + "'";

                            if (!Execute(Sql, out ErrorMSG))
                                alError.Add("DELETE stop_times => stop_id:" + STOP_ID + ",round_no:" + ROUND + ",stop_sequence:" + STA_NO + ",is_inbound:" + is_inbound + ",service_id:" + service_id + ",Error:" + ErrorMSG);

                            Sql = @"INSERT INTO [Google_Transit].[dbo].[stop_times]
                                           ([round_no]
                                           ,[stop_id]
                                           ,[line_id]
                                           ,[stop_sequence]
                                           ,[arrival_time]
                                           ,[departure_time]
                                           ,[is_inbound]
                                           ,[service_id]
                                           ,[stop_detail])
                                     VALUES
                                           ('@round_no'
                                           ,'@stop_id'
                                           ,'@line_id'
                                           ,'@stop_sequence'
                                           ,'@arrival_time'
                                           ,'@departure_time'
                                           ,@is_inbound
                                           ,'@service_id'
                                           ,'@stop_detail')";
                            Sql = Sql.Replace("@round_no", ROUND);
                            Sql = Sql.Replace("@stop_id", STOP_ID);
                            Sql = Sql.Replace("@line_id", tmp_line_id);
                            Sql = Sql.Replace("@stop_sequence", STA_NO);
                            Sql = Sql.Replace("@arrival_time", TO);
                            Sql = Sql.Replace("@departure_time", FROM);
                            Sql = Sql.Replace("@service_id", service_id);
                            Sql = Sql.Replace("@stop_detail", PathFile);
                            if (is_inbound)
                                Sql = Sql.Replace("@is_inbound", "1");
                            else
                                Sql = Sql.Replace("@is_inbound", "0");

                            if (!Execute(Sql, out ErrorMSG))
                                alError.Add("INSERT stop_times => stop_id:" + STOP_ID + ",round_no:" + ROUND + ",stop_sequence:" + STA_NO + ",is_inbound:" + is_inbound + ",service_id:" + service_id + ",Error:" + ErrorMSG);
                        }


                        if (cbxPrice.Checked)
                        {
                            for (int i = 0; i < dtPrice.Rows.Count; i++)
                            {
                                string STA_NO1 = "" + dtPrice.Rows[i]["STA_NO1"];
                                string STA_NO2 = "" + dtPrice.Rows[i]["STA_NO2"];
                                string PRICE = "" + dtPrice.Rows[i]["PRICE"];

                                DataRow[] DR = dtStop.Select("STA_NO = '" + STA_NO1 + "'");
                                string STOP_ID_1 = "" + DR[0]["STOP_ID"];
                                DR = dtStop.Select("STA_NO = '" + STA_NO2 + "'");
                                string STOP_ID_2 = "" + DR[0]["STOP_ID"];

                                string[] SP_PRICE = PRICE.Split(new char[] { ',' });


                                for (int k = 0; k < SP_PRICE.Length; k++)
                                {
                                    Sql = @"SELECT [fare_id]
                                          FROM [Google_Transit].[dbo].[fare_attributes]
                                        WHERE [price] = " + SP_PRICE[k] + " and [agency_id] = '" + Agency + "'";

                                    dtTmp = Retreive(Sql);
                                    string fare_id = "";

                                    if (dtTmp.Rows.Count != 0)
                                        fare_id = "" + dtTmp.Rows[0]["fare_id"];
                                    else
                                    {
                                        Sql = @"SELECT CONVERT(decimal(18,0),ISNULL(SUBSTRING(MAX(fare_id),7,LEN(MAX(fare_id))),0)) + 1 as new_id
                                            FROM [Google_Transit].[dbo].[fare_attributes]";

                                        dtTmp = Retreive(Sql);
                                        fare_id = "FA" + genId("" + dtTmp.Rows[0]["new_id"], 6);

                                        Sql = @"INSERT INTO [Google_Transit].[dbo].[fare_attributes]
                                                   ([agency_id]
                                                    ,[fare_id]
                                                   ,[price]
                                                   ,[currency_type]
                                                   ,[payment_method]
                                                   ,[transfers])
                                             VALUES
                                                   ('@agency_id','@fare_id'
                                                   ,'@price'
                                                   ,'THB'
                                                   ,'1'
                                                   ,'0')";
                                        Sql = Sql.Replace("@fare_id", fare_id);
                                        Sql = Sql.Replace("@agency_id", Agency);
                                        Sql = Sql.Replace("@price", SP_PRICE[k]);

                                        if (!Execute(Sql, out ErrorMSG))
                                            alError.Add("INSERT fare_attributes => fare_id:" + fare_id + ",price:" + SP_PRICE[k] + ",Error:" + ErrorMSG);
                                    }

                                    Sql = @"DELETE [Google_Transit].[dbo].[fare_rules]
                                    WHERE [line_id] = '" + tmp_line_id + "' AND [fare_id] = '" + fare_id + "' AND [origin_id] = '" + STA_NO1 + "' AND [destination_id] = '" + STA_NO2 + "' AND [service_id] = '" + service_id + "'";
                                    if (is_inbound)
                                        Sql = Sql + " AND is_inbound = 1";
                                    else
                                        Sql = Sql + " AND is_inbound = 0";


                                    if (!Execute(Sql, out ErrorMSG))
                                        alError.Add("DELETE fare_rules => line_id:" + tmp_line_id + ",stop_id_1:" + STOP_ID_1 + ",stop_id_2:" + STOP_ID_2 + ",fare_id:" + fare_id + ",Error:" + ErrorMSG);

                                    Sql = @"INSERT INTO [Google_Transit].[dbo].[fare_rules]
                                           ([fare_id]
                                           ,[line_id]
                                           ,[origin_id]
                                           ,[destination_id]
                                           ,[service_id]
                                           ,[is_inbound])
                                     VALUES
                                           ('@fare_id'
                                           ,'@line_id'
                                           ,'@origin_id'
                                           ,'@destination_id'
                                           ,'@service_id'
                                           ,'@is_inbound')";


                                    Sql = Sql.Replace("@fare_id", fare_id);
                                    Sql = Sql.Replace("@line_id", tmp_line_id);
                                    Sql = Sql.Replace("@origin_id", STOP_ID_1);
                                    Sql = Sql.Replace("@destination_id", STOP_ID_2);
                                    Sql = Sql.Replace("@service_id", service_id);
                                    if (is_inbound)
                                        Sql = Sql.Replace("@is_inbound", "1");
                                    else
                                        Sql = Sql.Replace("@is_inbound", "0");
                                    if (!Execute(Sql, out ErrorMSG))
                                        alError.Add("INSERT fare_rules => line_id:" + tmp_line_id + ",stop_id_1:" + STOP_ID_1 + ",stop_id_2:" + STOP_ID_2 + ",fare_id:" + fare_id + ",Error:" + ErrorMSG);

                                }



                            }
                        }
                        return true;
                    }
                    else
                    {
                        alError.Add("ReadExcel => Agency:" + Agency + ", Error:ไม่พบ Agency id นี้ในระบบ");
                        CloseExcel();
                        return false;
                    }                    
                }
                catch (Exception ex)
                {
                    alError.Add("ReadExcel => FilePath:" + PathFile + ", Error:" + ex.Message);
                    CloseExcel();
                    return false;
                }                
            }
            else
            {
                // Error Message อยู่ใน OpenExcel
                CloseExcel();
                return false;
            }
        }
        #endregion
        #region Function
        int strToInt(string str)
        {
            try
            {
                return int.Parse(str);
            }
            catch { return 0; }
        }
        double strToDouble(string str)
        {
            try
            {
                return double.Parse(str);
            }
            catch { return 0.0; }
        }
        public bool Execute(string sql, out string ErrorMsg)
        {
            #region Code
            try
            {
                SqlConnection conn = new SqlConnection(Connection);
                if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                conn.Open();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = sql;
                comm.CommandType = System.Data.CommandType.Text;
                comm.CommandTimeout = 600000;
                comm.Connection = conn;
                comm.ExecuteNonQuery();
                conn.Close();
                ErrorMsg = "";
                return true;
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
                return false;
            }
            #endregion
        }
        public System.Data.DataTable Retreive(string sql)
        {
            #region Code
            try
            {
                SqlConnection conn = new SqlConnection(Connection);

                if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                conn.Close();
                conn.Dispose();
                return dt;
            }
            catch (Exception ex)
            {               
                return null;
            }
            #endregion
        }
        public DataSet Retreives(string sql)
        {
            #region Code
            try
            {
                SqlConnection conn = new SqlConnection(Connection);

                if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds);
                conn.Close();
                conn.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }
        string genId(string str, int length)
        {
            int start = str.Length;
            for (int i = start; i < length; i++)
            {
                str = "0" + str;
            }
            return str;
        }       
        /// <summary>
        /// ใช้สำหรับนำเข้าข้อมูล KML โดยจะอัพเดทสองตาราง คือ Line,Stop
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="AgencyID"></param>
        /// <param name="route_type"></param>
        void fnLoad_KML(string FilePath, string AgencyID, string route_type)
        {
            try
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(FilePath);

                string xmlnsname = xDoc.ChildNodes[1].NamespaceURI.ToString();
                string strKml = File.ReadAllText(FilePath);
                strKml = strKml.Replace("xmlns=\"" + xmlnsname + "\"", "");

                XmlDocument xmlKml = new XmlDocument();
                xmlKml.LoadXml(strKml);

                string FileName = xmlKml.SelectSingleNode("//kml/Document/name").InnerText;
                XmlNodeList listFolder = xmlKml.SelectNodes("//kml/Document/Folder/Folder");
                if (listFolder.Count == 0)
                    listFolder = xmlKml.SelectNodes("//kml/Document/Folder");
                string FolderName = xmlKml.SelectSingleNode("//kml/Document/Folder/name").InnerText;
                              
                System.Data.DataTable dtTmp;

                string Sql = "";

                ArrayList listStopIn = new ArrayList();
                ArrayList listStopOut = new ArrayList();

                Random RD = new Random();
                string TmpLineRound = "TMP" + RD.Next(1000, 9999) + DateTime.Now.ToShortDateString();
                for (int i = 0; i < listFolder.Count; i++)
                {
                    string FolderType = listFolder[i].SelectSingleNode("./name").InnerText;
                 

                    XmlNodeList listPlacemark = listFolder[i].SelectNodes("./Placemark");
                    for (int j = 0; j < listPlacemark.Count; j++)
                    {
                        try
                        {
                            string shape = listPlacemark[j].SelectSingleNode("./LineString/coordinates").InnerText.Trim();
                            string[] SP_LST = shape.Split(new char[] { ' ' });
                            Sql = @"SELECT ISNULL(MAX([shape_id]),'') as [shape_id] FROM [Google_Transit].[dbo].[shapes]";
                            DataSet ds = Retreives(Sql);
                            string shape_id_max = ds.Tables[0].Rows[0][0].ToString().Trim();
                            string shape_new = "";
                            if (shape_id_max == "") shape_new = "SH00000001";
                            else
                            {
                                shape_new = shape_id_max.Replace("SH", "");
                                int new_id = strToInt(shape_new) + 1;
                                string tmp_id = "" + new_id;
                                int length_tmp = tmp_id.Length;
                                for (int z = length_tmp; z < 8; z++)
                                {
                                    tmp_id = "0" + tmp_id;
                                }
                                shape_new = "SH" + tmp_id;
                            }

                            if (AgencyID == "A007")
                            {
                                int k_id = 0;
                                for (int k = (SP_LST.Length-1); k >= 0; k--)
                                {
                                    Sql = @"INSERT INTO [Google_Transit].[dbo].[shapes]
                                                   ([shape_id]
                                                   ,[line_id]
                                                   ,[is_inbound]
                                                   ,[stop_id]
                                                   ,[stop_sequence]
                                                   ,[shape_pt_lat]
                                                   ,[shape_pt_lon]
                                                   ,[shape_dist_traveled]
                                                   ,[agency_id])
                                             VALUES
                                                   ('@shape_id'
                                                   ,'@line_id'
                                                   ,'@is_inbound'
                                                   ,''
                                                   ,'@stop_sequence'
                                                   ,'@shape_pt_lat'
                                                   ,'@shape_pt_lon'
                                                   ,''
                                                   ,'@agency_id')";

                                    Sql = Sql.Replace("@shape_id", shape_new);
                                    Sql = Sql.Replace("@agency_id", AgencyID);

                                    if (AgencyID == "A009")
                                    {
                                        FolderName = FolderName.Replace("ขบวน", "");
                                        Sql = Sql.Replace("@line_id", FolderName.Trim());
                                        Sql = Sql.Replace("@is_inbound", "1");
                                    }
                                    else
                                    {
                                        Sql = Sql.Replace("@line_id", FolderName.Trim());
                                        if (FolderType.ToUpper().Contains("INBOUND"))
                                            Sql = Sql.Replace("@is_inbound", "1");
                                        else
                                            Sql = Sql.Replace("@is_inbound", "0");
                                    }
                                    string[] SP_POINT = SP_LST[k].Split(new char[] { ',' });
                                    //100.5949736721657,13.57339133043166,0

                                    Sql = Sql.Replace("@stop_sequence", "" + (k_id + 1));
                                    Sql = Sql.Replace("@shape_pt_lat", SP_POINT[1]);
                                    Sql = Sql.Replace("@shape_pt_lon", SP_POINT[0]);

                                    string ErrorMSG = "";
                                    if (!Execute(Sql, out ErrorMSG))
                                        alError.Add("INSERT Feeds_Fail => Error:" + ErrorMSG);

                                    k_id++;
                                }
                            }
                            else
                            {
                                for (int k = 0; k < SP_LST.Length; k++)
                                {
                                    Sql = @"INSERT INTO [Google_Transit].[dbo].[shapes]
                                                   ([shape_id]
                                                   ,[line_id]
                                                   ,[is_inbound]
                                                   ,[stop_id]
                                                   ,[stop_sequence]
                                                   ,[shape_pt_lat]
                                                   ,[shape_pt_lon]
                                                   ,[shape_dist_traveled]
                                                   ,[agency_id])
                                             VALUES
                                                   ('@shape_id'
                                                   ,'@line_id'
                                                   ,'@is_inbound'
                                                   ,''
                                                   ,'@stop_sequence'
                                                   ,'@shape_pt_lat'
                                                   ,'@shape_pt_lon'
                                                   ,''
                                                   ,'@agency_id')";

                                    Sql = Sql.Replace("@shape_id", shape_new);
                                    Sql = Sql.Replace("@agency_id", AgencyID);


                                    if (AgencyID == "A009")
                                    {
                                        FolderName = FolderName.Replace("ขบวน", "");
                                        Sql = Sql.Replace("@line_id", FolderName.Trim());
                                        Sql = Sql.Replace("@is_inbound", "1");
                                    }
                                    else
                                    {
                                        Sql = Sql.Replace("@line_id", FolderName.Trim());
                                        if (FolderType.ToUpper().Contains("INBOUND"))
                                            Sql = Sql.Replace("@is_inbound", "1");
                                        else
                                            Sql = Sql.Replace("@is_inbound", "0");
                                    }
                                    string[] SP_POINT = SP_LST[k].Split(new char[] { ',' });
                                    //100.5949736721657,13.57339133043166,0

                                    Sql = Sql.Replace("@stop_sequence", "" + (k + 1));
                                    Sql = Sql.Replace("@shape_pt_lat", SP_POINT[1]);
                                    Sql = Sql.Replace("@shape_pt_lon", SP_POINT[0]);

                                    string ErrorMSG = "";
                                    if (!Execute(Sql, out ErrorMSG))
                                        alError.Add("INSERT Feeds_Fail => Error:" + ErrorMSG);
                                }
                            }
                            
                        }
                        catch{}
                    }
                }
            }
            catch (Exception ex)
            {
                alError.Add("FilePath => " + FilePath + ",Error:" + ex.Message);
            }
        }
        void genGTFS_Train()
        {
            DataSet ds = Retreives("EXEC [dbo].[spl_gen_A009]");
            System.Data.DataTable dt_routes = ds.Tables[0];
            System.Data.DataTable dt_stop = ds.Tables[1];
            System.Data.DataTable dt_shapes = ds.Tables[2];
            System.Data.DataTable dt_trips = ds.Tables[3];
            System.Data.DataTable dt_stop_times = ds.Tables[4];

            #region Routes
            string routes = File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/gtfs/routes.txt");
            for (int i = 0; i < dt_routes.Rows.Count; i++)
            {
                DataRow dr = dt_routes.Rows[i];
                //route_id,agency_id,route_short_name,route_long_name,route_desc,route_type,route_url,route_color,route_text_color
                string sLine = GV(dr["route_id"]) + "," + GV(dr["agency_id"]) + "," + GV(dr["route_short_name"]) + "," + GV(dr["route_long_name"])
                    + "," + GV(dr["route_desc"]) + "," + GV(dr["route_type"]) + "," + GV(dr["route_url"]) + "," + GV(dr["route_color"]) + "," + GV(dr["route_text_color"]);
                routes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/routes.txt", routes);
            #endregion
            #region Stops
            string stop = File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/gtfs/stops.txt");
            for (int i = 0; i < dt_stop.Rows.Count; i++)
            {
                DataRow dr = dt_stop.Rows[i];
                //stop_id,stop_name,stop_desc,stop_lat,stop_lon,stop_url
                string sLine = GV(dr["stop_id"]) + "," + GV(dr["stop_name"]) + "," + GV(dr["stop_desc"])
                    + "," + GV(dr["stop_lat"]) + "," + GV(dr["stop_lon"]) + "," + GV(dr["stop_url"]);
                stop += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stops.txt", stop);
            #endregion
            #region shapes
            string shapes = File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/gtfs/shapes.txt");
            for (int i = 0; i < dt_shapes.Rows.Count; i++)
            {
                DataRow dr = dt_shapes.Rows[i];
                //shape_id,shape_pt_lat,shape_pt_lon,shape_pt_sequence,shape_dist_traveled
                string sLine = GV(dr["shape_id"]) + "," + GV(dr["shape_pt_lat"]) + "," + GV(dr["shape_pt_lon"])
                    + "," + GV(dr["shape_pt_sequence"]) + "," + GV(dr["shape_dist_traveled"]);
                shapes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/shapes.txt", shapes);
            #endregion
            #region trips
            string trips = File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/gtfs/trips.txt");
            for (int i = 0; i < dt_trips.Rows.Count; i++)
            {
                DataRow dr = dt_trips.Rows[i];
                //route_id,service_id,trip_id,trip_headsign,direction_id,block_id,shape_id
                string sLine = GV(dr["route_id"]) + "," + GV(dr["service_id"]) + "," + GV(dr["trip_id"]) + "," + GV(dr["trip_headsign"])
                    + "," + GV(dr["direction_id"]) + "," + GV(dr["block_id"]) + "," + GV(dr["shape_id"]);
                trips += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/trips.txt", trips);
            #endregion
            #region stop_times
            //string stop_times = File.ReadAllText(Application.StartupPath + "/gtfs/stop_times.txt");
            //for (int i = 0; i < dt_stop_times.Rows.Count; i++)
            //{
            //    DataRow dr = dt_stop_times.Rows[i];
            //    //trip_id,arrival_time,departure_time,stop_id,stop_sequence,stop_headsign,pickup_type,drop_off_type,shape_dist_traveled
            //    string sLine = GV(dr["trip_id"]) + "," + GV(dr["arrival_time"]) + "," + GV(dr["departure_time"]) + "," + GV(dr["stop_id"])
            //        + "," + GV(dr["stop_sequence"]) + "," + GV(dr["stop_headsign"]) + "," + GV(dr["pickup_type"]) + "," + GV(dr["drop_off_type"])
            //        + "," + GV(dr["shape_dist_traveled"]);
            //    stop_times += Environment.NewLine + sLine;
            //}
            //File.WriteAllText(Application.StartupPath + "/output/stop_times.txt", stop_times);
            #endregion
        }
        void genGTFS_Train_Stoptime()
        {
            DataSet ds = Retreives("EXEC [dbo].[spl_gen_A009]");
            System.Data.DataTable dt_routes = ds.Tables[0];
            System.Data.DataTable dt_stop = ds.Tables[1];
            System.Data.DataTable dt_shapes = ds.Tables[2];
            System.Data.DataTable dt_trips = ds.Tables[3];
            System.Data.DataTable dt_stop_times = ds.Tables[4];
            #region stop_times
            string stop_times = "";// File.ReadAllText(Application.StartupPath + "/gtfs/stop_times.txt");
            for (int i = 0; i < dt_stop_times.Rows.Count; i++)
            {
                DataRow dr = dt_stop_times.Rows[i];
                //trip_id,arrival_time,departure_time,stop_id,stop_sequence,stop_headsign,pickup_type,drop_off_type,shape_dist_traveled
                string arrival_time = GV(dr["arrival_time"]);
                string departure_time = GV(dr["departure_time"]);
                if (arrival_time.Trim() == "-")
                    arrival_time = departure_time;
                if (departure_time.Trim() == "-")
                    departure_time = arrival_time;

                arrival_time = arrival_time + ":00";
                departure_time = departure_time + ":00";

                string sLine = GV(dr["trip_id"]) + "," + arrival_time + "," + departure_time + "," + GV(dr["stop_id"])
                    + "," + GV(dr["stop_sequence"]) + "," + GV(dr["stop_headsign"]) + "," + GV(dr["pickup_type"]) + "," + GV(dr["drop_off_type"])
                    + "," + GV(dr["shape_dist_traveled"]);
                stop_times += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stop_times.txt", stop_times);
            #endregion
        }
        string GV(object obj)
        {
            return "" + obj;
        }
        void genGTFS_Train_All()
        {
            DataSet ds = Retreives("EXEC [dbo].[spl_gen_A009]");
            System.Data.DataTable dt_routes = ds.Tables[0];
            System.Data.DataTable dt_stop = ds.Tables[1];
            System.Data.DataTable dt_shapes = ds.Tables[2];
            System.Data.DataTable dt_trips = ds.Tables[3];
            System.Data.DataTable dt_stop_times = ds.Tables[4];

            #region Routes
            string routes = "route_id,agency_id,route_short_name,route_long_name,route_desc,route_type,route_url,route_color,route_text_color";
            for (int i = 0; i < dt_routes.Rows.Count; i++)
            {
                DataRow dr = dt_routes.Rows[i];
                //route_id,agency_id,route_short_name,route_long_name,route_desc,route_type,route_url,route_color,route_text_color
                string sLine = GV(dr["route_id"]) + "," + GV(dr["agency_id"]) + "," + GV(dr["route_short_name"]) + "," + GV(dr["route_long_name"])
                    + "," + GV(dr["route_desc"]) + "," + GV(dr["route_type"]) + "," + GV(dr["route_url"]) + "," + GV(dr["route_color"]) + "," + GV(dr["route_text_color"]);
                routes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/routes.txt", routes);
            #endregion
            #region Stops
            string stop = "stop_id,stop_name,stop_desc,stop_lat,stop_lon,stop_url";
            for (int i = 0; i < dt_stop.Rows.Count; i++)
            {
                DataRow dr = dt_stop.Rows[i];
                //stop_id,stop_name,stop_desc,stop_lat,stop_lon,stop_url
                string sLine = GV(dr["stop_id"]) + "," + GV(dr["stop_name"]) + "," + GV(dr["stop_desc"])
                    + "," + GV(dr["stop_lat"]) + "," + GV(dr["stop_lon"]) + "," + GV(dr["stop_url"]);
                stop += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stops.txt", stop);
            #endregion
            #region shapes
            string shapes = "shape_id,shape_pt_lat,shape_pt_lon,shape_pt_sequence,shape_dist_traveled";
            for (int i = 0; i < dt_shapes.Rows.Count; i++)
            {
                DataRow dr = dt_shapes.Rows[i];
                //shape_id,shape_pt_lat,shape_pt_lon,shape_pt_sequence,shape_dist_traveled
                string sLine = GV(dr["shape_id"]) + "," + GV(dr["shape_pt_lat"]) + "," + GV(dr["shape_pt_lon"])
                    + "," + GV(dr["shape_pt_sequence"]) + "," + GV(dr["shape_dist_traveled"]);
                shapes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/shapes.txt", shapes);
            #endregion
            #region trips
            string trips = "route_id,service_id,trip_id,trip_headsign,direction_id,block_id,shape_id";
            for (int i = 0; i < dt_trips.Rows.Count; i++)
            {
                DataRow dr = dt_trips.Rows[i];
                //route_id,service_id,trip_id,trip_headsign,direction_id,block_id,shape_id
                string sLine = GV(dr["route_id"]) + "," + GV(dr["service_id"]) + "," + GV(dr["trip_id"]) + "," + GV(dr["trip_headsign"])
                    + "," + GV(dr["direction_id"]) + "," + GV(dr["block_id"]) + "," + GV(dr["shape_id"]);
                trips += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/trips.txt", trips);
            #endregion
            #region stop_times
            string stop_times = "trip_id,arrival_time,departure_time,stop_id,stop_sequence,stop_headsign,pickup_type,drop_off_type,shape_dist_traveled";// File.ReadAllText(Application.StartupPath + "/gtfs/stop_times.txt");
            for (int i = 0; i < dt_stop_times.Rows.Count; i++)
            {
                DataRow dr = dt_stop_times.Rows[i];
                //trip_id,arrival_time,departure_time,stop_id,stop_sequence,stop_headsign,pickup_type,drop_off_type,shape_dist_traveled
                string arrival_time = GV(dr["arrival_time"]);
                string departure_time = GV(dr["departure_time"]);
                if (arrival_time.Trim() == "-")
                    arrival_time = departure_time;
                if (departure_time.Trim() == "-")
                    departure_time = arrival_time;

                arrival_time = arrival_time + ":00";
                departure_time = departure_time + ":00";

                string sLine = GV(dr["trip_id"]) + "," + arrival_time + "," + departure_time + "," + GV(dr["stop_id"])
                    + "," + GV(dr["stop_sequence"]) + "," + GV(dr["stop_headsign"]) + "," + GV(dr["pickup_type"]) + "," + GV(dr["drop_off_type"])
                    + "," + GV(dr["shape_dist_traveled"]);
                stop_times += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stop_times.txt", stop_times);
            #endregion
        }
        void genGTFS_By_Agency(string agency_id)
        {
            string MsgError = "";
            Execute("EXEC [dbo].[spl_gen_gtfs]", out MsgError);
            DataSet ds = Retreives("EXEC [dbo].[spl_gen_A009]");

            #region Agency
            string Sql = @"SELECT [agency_id]
                                  ,[agency_name]
                                  ,[agency_detail]
                                  ,[agency_url]
                                  ,[agency_timezone]
                                  ,[agency_lang]
                                  ,[agency_phone]
                                  ,[agency_fare_url]
                                  ,[agency_email]
                                  ,[RowId]
                              FROM [Google_Transit].[dbo].[agency]
                            WHERE agency_id = '" + agency_id + "'";
            System.Data.DataTable dt_agency = Retreive(Sql);

            string Agency = "agency_id,agency_name,agency_url,agency_timezone,agency_lang,agency_phone,agency_fare_url,agency_email";
            for (int i = 0; i < dt_agency.Rows.Count; i++)
            {
                DataRow dr = dt_agency.Rows[i];
                string sLine = GV(dr["agency_id"]) + "," + GV(dr["agency_name"]) + "," + GV(dr["agency_url"]) + "," + GV(dr["agency_timezone"]) + "," + GV(dr["agency_lang"]) + "," + GV(dr["agency_phone"]) + "," + GV(dr["agency_fare_url"]) + "," + GV(dr["agency_email"]);
                Agency += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/agency.txt", Agency);
            #endregion
            #region calendar
            Sql = @"SELECT [service_id]
                          ,[monday]
                          ,[tuesday]
                          ,[wednesday]
                          ,[thursday]
                          ,[friday]
                          ,[saturday]
                          ,[sunday]
                          ,[start_date]
                          ,[end_date]
                      FROM [Google_Transit].[dbo].[v_gtfs_calendar]";
            System.Data.DataTable dt_calendar = Retreive(Sql);

            string Calendar = "service_id,monday,tuesday,wednesday,thursday,friday,saturday,sunday,start_date,end_date";
            for (int i = 0; i < dt_calendar.Rows.Count; i++)
            {
                DataRow dr = dt_calendar.Rows[i];
                string sLine = GV(dr["service_id"]) + ","
                    + GV(dr["monday"]) + ","
                    + GV(dr["tuesday"]) + ","
                    + GV(dr["wednesday"]) + ","
                    + GV(dr["thursday"]) + ","
                    + GV(dr["friday"]) + ","
                    + GV(dr["saturday"]) + ","
                    + GV(dr["sunday"]) + ","
                    + GV(dr["start_date"]) + ","
                    + GV(dr["end_date"]);
                Calendar += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/calendar.txt", Calendar);
            #endregion
            #region fare_attributes
            Sql = @"SELECT [fare_id]
                          ,[price]
                          ,[currency_type]
                          ,[payment_method]
                          ,[transfers]
                      FROM [Google_Transit].[dbo].[v_gtfs_fare_attributes]";
            System.Data.DataTable dt_fare_attributes = Retreive(Sql);

            string fare_attributes = "fare_id,price,currency_type,payment_method,transfers";
            for (int i = 0; i < dt_fare_attributes.Rows.Count; i++)
            {
                DataRow dr = dt_fare_attributes.Rows[i];
                string sLine = GV(dr["fare_id"]) + ","
                    + GV(dr["price"]) + ","
                    + GV(dr["currency_type"]) + ","
                    + GV(dr["payment_method"]) + ","
                    + GV(dr["transfers"]);
                fare_attributes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/fare_attributes.txt", fare_attributes);
            #endregion
            #region fare_rules
            Sql = @"SELECT [fare_id]
                          ,[route_id]
                          ,[origin_id]
                          ,[destination_id]
                      FROM [Google_Transit].[dbo].[v_gtfs_fare_rules]
                     WHERE [route_id] in(
	                      SELECT [route_id]      
	                      FROM [Google_Transit].[dbo].[v_gtfs_routes]
	                      WHERE agency_id = '" + agency_id + "')";
            System.Data.DataTable dt_fare_rules = Retreive(Sql);

            string fare_rules = "fare_id,route_id,origin_id,destination_id";
            for (int i = 0; i < dt_fare_rules.Rows.Count; i++)
            {
                DataRow dr = dt_fare_rules.Rows[i];
                string sLine = GV(dr["fare_id"]) + ","
                    + GV(dr["route_id"]) + ","
                    + GV(dr["origin_id"]) + ","
                    + GV(dr["destination_id"]);
                fare_rules += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/fare_rules.txt", fare_rules);
            #endregion
            #region routes
            Sql = @"SELECT [route_id]
                          ,[agency_id]
                          ,[route_short_name]
                          ,[route_long_name]
                          ,[route_desc]
                          ,[route_type]
                          ,[route_url]
                          ,[route_color]
                          ,[route_text_color]
                      FROM [Google_Transit].[dbo].[v_gtfs_routes]
                    WHERE agency_id = '" + agency_id + "'";
            System.Data.DataTable dt_routes = Retreive(Sql);

            string routes = "route_id,agency_id,route_short_name,route_long_name,route_desc,route_type,route_url,route_color,route_text_color";
            for (int i = 0; i < dt_routes.Rows.Count; i++)
            {
                DataRow dr = dt_routes.Rows[i];
                string sLine = GV(dr["route_id"]) + ","
                    + GV(dr["agency_id"]) + ","
                    + GV(dr["route_short_name"]) + ","
                    + GV(dr["route_long_name"]) + ","
                    + GV(dr["route_desc"]) + ","
                    + GV(dr["route_type"]) + ","
                    + GV(dr["route_url"]) + ","
                    + GV(dr["route_color"]) + ","
                    + GV(dr["route_text_color"]);
                routes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/routes.txt", routes);
            #endregion
            #region shapes
            Sql = @"SELECT [shape_id]
                          ,[shape_pt_lat]
                          ,[shape_pt_lon]
                          ,[shape_pt_sequence]
                          ,[shape_dist_traveled]
                      FROM [Google_Transit].[dbo].[v_gtfs_shapes]
                    WHERE [shape_id] in (
	                    SELECT [shape_id]
	                      FROM [Google_Transit].[dbo].[v_gtfs_trips]
	                      WHERE [route_id] in(
		                      SELECT [route_id]      
		                      FROM [Google_Transit].[dbo].[v_gtfs_routes]
		                      WHERE agency_id = '" + agency_id + @"'
	                      )
	                    GROUP BY [shape_id]
                      )";
            System.Data.DataTable dt_shapes = Retreive(Sql);

            string shapes = "shape_id,shape_pt_lat,shape_pt_lon,shape_pt_sequence,shape_dist_traveled";
            for (int i = 0; i < dt_shapes.Rows.Count; i++)
            {
                DataRow dr = dt_shapes.Rows[i];
                string sLine = GV(dr["shape_id"]) + ","
                    + GV(dr["shape_pt_lat"]) + ","
                    + GV(dr["shape_pt_lon"]) + ","
                    + GV(dr["shape_pt_sequence"]) + ","
                    + GV(dr["shape_dist_traveled"]);
                shapes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/shapes.txt", shapes);
            #endregion
            #region stops
            Sql = @"SELECT [stop_id]
                          ,[stop_name]
                          ,[stop_desc]
                          ,[stop_lat]
                          ,[stop_lon]
                          ,[zone_id]
                          ,[stop_url]
                      FROM [Google_Transit].[dbo].[v_gtfs_stops]";
            System.Data.DataTable dt_stops = Retreive(Sql);

            string stops = "stop_id,stop_name,stop_desc,stop_lat,stop_lon,zone_id,stop_url";
            for (int i = 0; i < dt_stops.Rows.Count; i++)
            {
                DataRow dr = dt_stops.Rows[i];
                string sLine = GV(dr["stop_id"]) + ","
                    + GV(dr["stop_name"]) + ","
                    + GV(dr["stop_desc"]) + ","
                    + GV(dr["stop_lat"]) + ","
                    + GV(dr["stop_lon"]) + ","
                    + GV(dr["zone_id"]) + ","
                    + GV(dr["stop_url"]);
                stops += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stops.txt", stops);
            #endregion
            #region stop_times
            Sql = @"SELECT [trip_id]
                          ,[arrival_time]
                          ,[departure_time]
                          ,[stop_id]
                          ,[stop_sequence]
                      FROM [Google_Transit].[dbo].[v_gtfs_stop_times]
                    WHERE trip_id in(
	                    SELECT [trip_id]
	                      FROM [Google_Transit].[dbo].[v_gtfs_trips]
	                      WHERE [route_id] in(
		                      SELECT [route_id]      
		                      FROM [Google_Transit].[dbo].[v_gtfs_routes]
		                      WHERE agency_id = '" + agency_id + @"'
	                      )
	                    GROUP BY [trip_id]
                      )";
            System.Data.DataTable dt_stop_times = Retreive(Sql);

            string stop_times = "trip_id,arrival_time,departure_time,stop_id,stop_sequence";
            for (int i = 0; i < dt_stop_times.Rows.Count; i++)
            {
                DataRow dr = dt_stop_times.Rows[i];
                string sLine = GV(dr["trip_id"]) + ","
                    + GV(dr["arrival_time"]) + ","
                    + GV(dr["departure_time"]) + ","
                    + GV(dr["stop_id"]) + ","
                    + GV(dr["stop_sequence"]);
                stop_times += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stop_times.txt", stop_times);
            #endregion
            #region trips
            Sql = @"SELECT [route_id]
                          ,[service_id]
                          ,[trip_id]
                          ,[direction_id]
                          ,[shape_id]
                      FROM [Google_Transit].[dbo].[v_gtfs_trips]
                        WHERE [route_id] in(
	                      SELECT [route_id]      
	                      FROM [Google_Transit].[dbo].[v_gtfs_routes]
	                      WHERE agency_id = '" + agency_id + "')";
            System.Data.DataTable dt_trips = Retreive(Sql);

            string trips = "route_id,service_id,trip_id,direction_id,shape_id";
            for (int i = 0; i < dt_trips.Rows.Count; i++)
            {
                DataRow dr = dt_trips.Rows[i];
                string sLine = GV(dr["route_id"]) + ","
                    + GV(dr["service_id"]) + ","
                    + GV(dr["trip_id"]) + ","
                    + GV(dr["direction_id"]) + ","
                    + GV(dr["shape_id"]);
                trips += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/trips.txt", trips);
            #endregion

            MsgOutput = "สร้างไฟล์ GTFS สำเร็จ";
        }
        void genGTFS_All()
        {
            string MsgError = "";
            Execute("EXEC [dbo].[spl_gen_gtfs]", out MsgError);
            DataSet ds = Retreives("EXEC [dbo].[spl_gen_A009]");


            #region Agency
            string Sql = @"SELECT [agency_id]
                                  ,[agency_name]
                                  ,[agency_detail]
                                  ,[agency_url]
                                  ,[agency_timezone]
                                  ,[agency_lang]
                                  ,[agency_phone]
                                  ,[agency_fare_url]
                                  ,[agency_email]
                                  ,[RowId]
                              FROM [Google_Transit].[dbo].[agency]";
            System.Data.DataTable dt_agency = Retreive(Sql);

            string Agency = "agency_id,agency_name,agency_url,agency_timezone,agency_lang,agency_phone,agency_fare_url,agency_email";
            for (int i = 0; i < dt_agency.Rows.Count; i++)
            {
                DataRow dr = dt_agency.Rows[i];
                string sLine = GV(dr["agency_id"]) + "," + GV(dr["agency_name"]) + "," + GV(dr["agency_url"]) + "," + GV(dr["agency_timezone"]) + "," + GV(dr["agency_lang"]) + "," + GV(dr["agency_phone"]) + "," + GV(dr["agency_fare_url"]) + "," + GV(dr["agency_email"]);
                Agency += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/agency.txt", Agency);
            #endregion
            #region calendar
            Sql = @"SELECT [service_id]
                          ,[monday]
                          ,[tuesday]
                          ,[wednesday]
                          ,[thursday]
                          ,[friday]
                          ,[saturday]
                          ,[sunday]
                          ,[start_date]
                          ,[end_date]
                      FROM [Google_Transit].[dbo].[v_gtfs_calendar]";
            System.Data.DataTable dt_calendar = Retreive(Sql);

            string Calendar = "service_id,monday,tuesday,wednesday,thursday,friday,saturday,sunday,start_date,end_date";
            for (int i = 0; i < dt_calendar.Rows.Count; i++)
            {
                DataRow dr = dt_calendar.Rows[i];
                string sLine = GV(dr["service_id"]) + ","
                    + GV(dr["monday"]) + ","
                    + GV(dr["tuesday"]) + ","
                    + GV(dr["wednesday"]) + ","
                    + GV(dr["thursday"]) + ","
                    + GV(dr["friday"]) + ","
                    + GV(dr["saturday"]) + ","
                    + GV(dr["sunday"]) + ","
                    + GV(dr["start_date"]) + ","
                    + GV(dr["end_date"]);
                Calendar += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/calendar.txt", Calendar);
            #endregion
            #region fare_attributes
            Sql = @"SELECT [agency_id]
                          ,[fare_id]
                          ,[price]
                          ,[currency_type]
                          ,[payment_method]
                          ,[transfers]
                      FROM [Google_Transit].[dbo].[v_gtfs_fare_attributes]";
            System.Data.DataTable dt_fare_attributes = Retreive(Sql);

            string fare_attributes = "agency_id,fare_id,price,currency_type,payment_method,transfers";
            for (int i = 0; i < dt_fare_attributes.Rows.Count; i++)
            {
                DataRow dr = dt_fare_attributes.Rows[i];
                string sLine = GV(dr["agency_id"]) + ","
                    + GV(dr["fare_id"]) + ","
                    + GV(dr["price"]) + ","
                    + GV(dr["currency_type"]) + ","
                    + GV(dr["payment_method"]) + ","
                    + GV(dr["transfers"]);
                fare_attributes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/fare_attributes.txt", fare_attributes);
            #endregion
            #region fare_rules
            Sql = @"SELECT [fare_id]
                          ,[route_id]
                          ,[origin_id]
                          ,[destination_id]
                      FROM [Google_Transit].[dbo].[v_gtfs_fare_rules]";
            System.Data.DataTable dt_fare_rules = Retreive(Sql);

            string fare_rules = "fare_id,route_id,origin_id,destination_id";
            for (int i = 0; i < dt_fare_rules.Rows.Count; i++)
            {
                DataRow dr = dt_fare_rules.Rows[i];
                string sLine = GV(dr["fare_id"]) + ","
                    + GV(dr["route_id"]) + ","
                    + GV(dr["origin_id"]) + ","
                    + GV(dr["destination_id"]);
                fare_rules += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/fare_rules.txt", fare_rules);
            #endregion
            #region routes
            Sql = @"SELECT [route_id]
                          ,[agency_id]
                          ,[route_short_name]
                          ,[route_long_name]
                          ,[route_desc]
                          ,[route_type]
                          ,[route_url]
                          ,[route_color]
                          ,[route_text_color]
                      FROM [Google_Transit].[dbo].[v_gtfs_routes]";
            System.Data.DataTable dt_routes = Retreive(Sql);

            string routes = "route_id,agency_id,route_short_name,route_long_name,route_desc,route_type,route_url,route_color,route_text_color";
            for (int i = 0; i < dt_routes.Rows.Count; i++)
            {
                DataRow dr = dt_routes.Rows[i];
                string sLine = GV(dr["route_id"]) + ","
                    + GV(dr["agency_id"]) + ","
                    + GV(dr["route_short_name"]) + ","
                    + GV(dr["route_long_name"]) + ","
                    + GV(dr["route_desc"]) + ","
                    + GV(dr["route_type"]) + ","
                    + GV(dr["route_url"]) + ","
                    + GV(dr["route_color"]) + ","
                    + GV(dr["route_text_color"]);
                routes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/routes.txt", routes);
            #endregion
            #region shapes
            Sql = @"SELECT [shape_id]
                          ,[shape_pt_lat]
                          ,[shape_pt_lon]
                          ,[shape_pt_sequence]
                          ,[shape_dist_traveled]
                      FROM [Google_Transit].[dbo].[v_gtfs_shapes]";
            System.Data.DataTable dt_shapes = Retreive(Sql);

            string shapes = "shape_id,shape_pt_lat,shape_pt_lon,shape_pt_sequence,shape_dist_traveled";
            for (int i = 0; i < dt_shapes.Rows.Count; i++)
            {
                DataRow dr = dt_shapes.Rows[i];
                string sLine = GV(dr["shape_id"]) + ","
                    + GV(dr["shape_pt_lat"]) + ","
                    + GV(dr["shape_pt_lon"]) + ","
                    + GV(dr["shape_pt_sequence"]) + ","
                    + GV(dr["shape_dist_traveled"]);
                shapes += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/shapes.txt", shapes);
            #endregion
            #region stops
            Sql = @"SELECT [stop_id]
                          ,[stop_name]
                          ,[stop_desc]
                          ,[stop_lat]
                          ,[stop_lon]
                          ,[zone_id]
                          ,[stop_url]
                      FROM [Google_Transit].[dbo].[v_gtfs_stops]
                    WHERE stop_id in(
                    SELECT [stop_id]
                      FROM [Google_Transit].[dbo].[stop_times]
                      GROUP BY [stop_id])";
            System.Data.DataTable dt_stops = Retreive(Sql);

            string stops = "stop_id,stop_name,stop_desc,stop_lat,stop_lon,zone_id,stop_url";
            for (int i = 0; i < dt_stops.Rows.Count; i++)
            {
                DataRow dr = dt_stops.Rows[i];
                string sLine = GV(dr["stop_id"]) + ","
                    + GV(dr["stop_name"]) + ","
                    + GV(dr["stop_desc"]) + ","
                    + GV(dr["stop_lat"]) + ","
                    + GV(dr["stop_lon"]) + ","
                    + GV(dr["zone_id"]) + ","
                    + GV(dr["stop_url"]);
                stops += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stops.txt", stops);
            #endregion
            #region stop_times
            //Sql = @"SELECT [trip_id]
            //              ,[arrival_time]
            //              ,[departure_time]
            //              ,[stop_id]
            //              ,[stop_sequence]
            //          FROM [Google_Transit].[dbo].[v_gtfs_stop_times]";
            Sql = @"SELECT [trip_id]+','+[arrival_time]+','+[departure_time]+','+[stop_id]+','+convert(varchar(200),[stop_sequence]) as [data]
                     FROM [Google_Transit].[dbo].[v_gtfs_stop_times]
                     ORDER BY [trip_id],[stop_sequence]";
            System.Data.DataTable dt_stop_times = Retreive(Sql);

            string stop_times = "trip_id,arrival_time,departure_time,stop_id,stop_sequence";
            for (int i = 0; i < dt_stop_times.Rows.Count; i++)
            {
                //DataRow dr = dt_stop_times.Rows[i];
                //string sLine = GV(dr["trip_id"]) + ","
                //    + GV(dr["arrival_time"]) + ","
                //    + GV(dr["departure_time"]) + ","
                //    + GV(dr["stop_id"]) + ","
                //    + GV(dr["stop_sequence"]);
                //stop_times += Environment.NewLine + sLine;
                stop_times += Environment.NewLine + dt_stop_times.Rows[i]["data"].ToString();
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/stop_times.txt", stop_times);
            #endregion
            #region trips
            Sql = @"SELECT [route_id]
                          ,[service_id]
                          ,[trip_id]
                          ,[direction_id]
                          ,[shape_id]
                      FROM [Google_Transit].[dbo].[v_gtfs_trips]";
            System.Data.DataTable dt_trips = Retreive(Sql);

            string trips = "route_id,service_id,trip_id,direction_id,shape_id";
            for (int i = 0; i < dt_trips.Rows.Count; i++)
            {
                DataRow dr = dt_trips.Rows[i];
                string sLine = GV(dr["route_id"]) + ","
                    + GV(dr["service_id"]) + ","
                    + GV(dr["trip_id"]) + ","
                    + GV(dr["direction_id"]) + ","
                    + GV(dr["shape_id"]);
                trips += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/trips.txt", trips);
            #endregion
            #region frequencies
            Sql = @"SELECT [trip_id]
		                    ,'04:50:00' as [start_time]
		                    ,'22:00:00' as [end_time]
		                    ,15*60 as headway_secs
                      FROM [Google_Transit].[dbo].[v_gtfs_trips]
                      GROUP BY [trip_id]
                      ORDER BY  [trip_id]";
            System.Data.DataTable dt_frequencies = Retreive(Sql);

            string frequencies = "trip_id,start_time,end_time,headway_secs";
            for (int i = 0; i < dt_frequencies.Rows.Count; i++)
            {
                DataRow dr = dt_frequencies.Rows[i];
                string sLine = GV(dr["trip_id"]) + ","
                    + GV(dr["start_time"]) + ","
                    + GV(dr["end_time"]) + ","
                    + GV(dr["headway_secs"]);
                frequencies += Environment.NewLine + sLine;
            }
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output/frequencies.txt", frequencies);
            #endregion

            MsgOutput = "สร้างไฟล์ GTFS สำเร็จ";
        }
        private DataRow AddRow(System.Data.DataTable DT, string Code, string Text)
        {
            DataRow DR = DT.NewRow();
            DR.BeginEdit();
            DR["Code"] = Code;
            DR["Text"] = Text;
            DR.EndEdit();
            return DR;
        }
        #endregion
        #region google api
        private string funChkTime_from_api_distance(string A, string B)
        {
            DataSet ds = new DataSet();
            ds.ReadXml("http://maps.google.com/maps/api/directions/xml?origin=" + B + "&destination=" + A + "&sensor=false");
            //string times = ds.Tables[7].Rows[0]["text"].ToString();
            System.Data.DataTable dt = null; 
            try
            {
                //หาระยะทาง
                dt = ds.Tables["distance"];
                int out_distance = 9999999;
                string text_distance = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    int distance = strToInt(dt.Rows[i]["value"].ToString());
                    if (distance < out_distance)
                    {
                        out_distance = distance;
                        text_distance = "" + dt.Rows[i]["text"];
                    }
                }
                //หาเวลาที่ใช้
                dt = ds.Tables["duration"];
                int out_time = 99999999;
                string text_time = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    int time = strToInt(dt.Rows[i]["value"].ToString());
                    if (time < out_time)
                    {
                        out_time = time;
                        text_time = "" + dt.Rows[i]["text"];
                    }
                }
                //หารายละเอียดเส้นทาง
                string route_txt = "" + ds.Tables["route"].Rows[0]["summary"];
                string start_address = "" + ds.Tables["leg"].Rows[0]["start_address"];
                string end_address = "" + ds.Tables["leg"].Rows[0]["end_address"];

                if (out_distance == 9999999)
                    return "";
                else
                    return "" + out_distance;
            }
            catch {  }
            //Thread.Sleep(3000);
            return "";
        }
        #endregion

        ArrayList alError = new ArrayList();
        string TimeZone = "Asia/Bangkok,th";
        string Connection = "Data Source=localhost; uid=sa; pwd=nextwaver; Initial Catalog=Google_Transit;";
        
        public frmMain()
        {
            InitializeComponent();
        }


        #region Event Function
        private void btnImportBus_Click(object sender, EventArgs e)
        { 
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                FileInfo[] fileInfo = dirInfo.GetFiles("*.kml");
                for (int i = 0; i < fileInfo.Length; i++)
                {
                    fnLoad_KML(fileInfo[i].FullName, "A003", "3");
                }
                string ErrorMSG = "";
                for (int i = 0; i < alError.Count; i++)
                {
                    ErrorMSG += alError[i] + Environment.NewLine;
                }
            }
        }
        private void btnTrain_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                FileInfo[] fileInfo = dirInfo.GetFiles("*.kml");
                for (int i = 0; i < fileInfo.Length; i++)
                {
                    fnLoad_KML(fileInfo[i].FullName, "A009", "2");
                }

                string ErrorMSG = "";
                for (int i = 0; i < alError.Count; i++)
                {
                    ErrorMSG += alError[i] + Environment.NewLine;
                }
            }
        }        
        private void button1_Click(object sender, EventArgs e)
        {
            genGTFS_Train();
            MessageBox.Show("OK");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            genGTFS_Train_Stoptime();
            MessageBox.Show("OK");
        }       
        private void button3_Click(object sender, EventArgs e)
        {
            genGTFS_Train_All();
            MessageBox.Show("OK");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                FileInfo[] fileInfo = dirInfo.GetFiles("*.kml");
                for (int i = 0; i < fileInfo.Length; i++)
                {
                    fnLoad_KML(fileInfo[i].FullName, "A005", "4");
                }

                string ErrorMSG = "";
                for (int i = 0; i < alError.Count; i++)
                {
                    ErrorMSG += alError[i] + Environment.NewLine;
                }
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                alError = new ArrayList();
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                fileInfo = dirInfo.GetFiles("*.xls");
                if (fileInfo.Length == 0)
                {
                    fileInfo = dirInfo.GetFiles("*.xlsx");
                }
                isINBOUND = true;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(RunLoadExcel);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;

                button5.Text = "2. INBOUND นำเข้าข้อมูลจาก Template (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
            }            
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string Sql = @"SELECT [line],[detail]
                          FROM [Google_Transit].[dbo].[A009]
                          GROUP BY [line],[detail]";
            System.Data.DataTable DT = Retreive(Sql);
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string line = "" + DT.Rows[i]["line"];
                string detail = "" + DT.Rows[i]["detail"];
                Sql = @"SELECT [line]
                               ,[detail]
                               ,[seq]
                               ,[stop_name_th]
                               ,[stop_name_en]
                               ,[start_time]
                               ,[end_time]
                               ,[find_name]
                               ,[stop_lat]
                               ,[stop_lon]
                         FROM [Google_Transit].[dbo].[A009]
                         WHERE [line] = '@line'
                         ORDER BY line, seq";
                Sql = Sql.Replace("@line", line);
                System.Data.DataTable dtTemp = Retreive(Sql);

                string strData = "<Folder><name>1</name><open>1</open><Folder><name>INBOUND</name>";
                for (int j = 0; j < dtTemp.Rows.Count; j++)
                {
                    string strTmp = @"<Placemark>
				                        <name>[ INBOUND ]@name</name>
				                        <open>1</open>
				                        <LookAt>
					                        <longitude>@longitude</longitude>
					                        <latitude>@latitude</latitude>
					                        <altitude>0</altitude>
					                        <heading>0.006033119647067864</heading>
					                        <tilt>0</tilt>
					                        <range>200</range>
					                        <altitudeMode>relativeToGround</altitudeMode>
				                        </LookAt>
				                        <styleUrl>#m_ylw-pushpin10</styleUrl>
				                        <Point>
					                        <coordinates>@longitude,@latitude,0</coordinates>
				                        </Point>
			                        </Placemark>";
                    strTmp = strTmp.Replace("@name", "" + dtTemp.Rows[j]["stop_name_th"]);
                    strTmp = strTmp.Replace("@longitude", "" + dtTemp.Rows[j]["stop_lon"]);
                    strTmp = strTmp.Replace("@latitude", "" + dtTemp.Rows[j]["stop_lat"]);

                    strData += strTmp;
                }
                strData += "</Folder></Folder>";

                string outputXML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
                outputXML += "<kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:kml=\"http://www.opengis.net/kml/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\">";
                outputXML += "<Document>";
                outputXML += "<name>" + line + ".kml</name>";
                outputXML += File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/config/style_kml.txt");
                outputXML += strData;
                outputXML += "</Document>";
                outputXML += "</kml>";


                File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output_kml/" + line + " " + detail + ".kml", outputXML);
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            string Sql = @"SELECT [line_code]      
                                  ,[inbound_detail]
                                  ,[outbound_detail]      
                              FROM [Google_Transit].[dbo].[line]
                              WHERE agency_id = 'A003'
                              GROUP BY [line_code]      
                                  ,[inbound_detail]
                                  ,[outbound_detail]";
            System.Data.DataTable DT = Retreive(Sql);
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string line_code = "" + DT.Rows[i]["line_code"];
                string inbound_detail = "" + DT.Rows[i]["inbound_detail"];
                string outbound_detail = "" + DT.Rows[i]["outbound_detail"];
                
                #region Query INBOUND
                Sql = @"DECLARE @IS_INBOUND int = 1
                        DECLARE @LINE varchar(200) = '" + line_code + @"'
                        DECLARE @INBOUND_TXT varchar(200) = 'INBOUND'                        

                        DECLARE @TmpWay Table(
	                        [route_id] varchar(100),
	                        [stop_id] varchar(100),
	                        [stop_name_th] varchar(500),
	                        [seq] int
                        )
                        INSERT INTO @TmpWay
                        SELECT w.[route_id]
                              ,w.[stop_id]
                              ,c.stop_name_th
                              ,w.[seq]
                          FROM [Google_Transit].[dbo].[way] w   
                          LEFT JOIN [Google_Transit].[dbo].[stop] c
	                        ON w.stop_id = c.stop_id
                          WHERE [route_id] in (SELECT [route_id]     
					                          FROM [Google_Transit].[dbo].[routes]
					                          WHERE [route_code] = @LINE 
					                          AND [is_inbound] = @IS_INBOUND)
                        ORDER by seq

                        SELECT A.[route_id]
                              ,A.[stop_id]
                              ,A.[stop_name_th]
                              ,B._div
                              ,B.DIV
                              ,B._time
                              ,B.TIME
                              ,A.[seq]
                        FROM ( SELECT A1.route_id,
			                        A1.stop_id,
			                        A1.stop_name_th,
			                        A2.stop_id as stop_id_2,
			                        A2.stop_name_th as stop_name_th_2,
			                        A1.seq
		                        FROM @TmpWay A1
		                        LEFT JOIN @TmpWay A2
			                        ON A1.seq+1 = (A2.seq)
		                        WHERE A2.stop_id is not null
                        ) A
                        LEFT JOIN [Google_Transit].[dbo].BUSSTOP_TIME_REVERT B
	                        ON A.stop_id = B.stop_id 
	                        AND A.stop_id_2 = B.stop_id_2
	                        AND B.BOUND = @INBOUND_TXT";
                #endregion
                #region INBOUND
                System.Data.DataTable DT_INBOUND = Retreive(Sql);                
                if (OpenExcel(System.Windows.Forms.Application.StartupPath+"/config/Template.xls"))
                {
                    SheetData.get_Range("B3", missing).Value2 = line_code;
                    SheetData.get_Range("B4", missing).Value2 = inbound_detail;
                    SheetData.get_Range("F5", missing).Value2 = "" + DT_INBOUND.Rows.Count;
                    int StartIndex = 12;
                    for (int j = 0; j < DT_INBOUND.Rows.Count; j++)
                    {
                        string seq = "" + DT_INBOUND.Rows[j]["seq"];
                        string stop_Name_th = "" + DT_INBOUND.Rows[j]["stop_Name_th"];
                        string _div = "" + DT_INBOUND.Rows[j]["_div"];
                        string DIV = "" + DT_INBOUND.Rows[j]["DIV"];
                        string _time = "" + DT_INBOUND.Rows[j]["_time"];
                        string TIME = "" + DT_INBOUND.Rows[j]["TIME"];

                        DIV = DIV.Replace("km", "");
                        SheetData.get_Range("A" + (StartIndex + j), missing).Value2 = seq;
                        SheetData.get_Range("B" + (StartIndex + j), missing).Value2 = stop_Name_th;
                        SheetData.get_Range("D" + (StartIndex + j), missing).Value2 = DIV.Trim();
                        SheetData.get_Range("F" + (StartIndex + j), missing).Value2 = _time;                        
                    }
                    SaveExcel(System.Windows.Forms.Application.StartupPath + "/output_excel/" + line_code + "(INBOUND).xls");
                }
                CloseExcel();
                #endregion

                #region Query OUTBOUND
                Sql = @"DECLARE @IS_INBOUND int = 0
                        DECLARE @LINE varchar(200) = '" + line_code + @"'
                        DECLARE @INBOUND_TXT varchar(200) = 'OUTBOUND'                        

                        DECLARE @TmpWay Table(
	                        [route_id] varchar(100),
	                        [stop_id] varchar(100),
	                        [stop_name_th] varchar(500),
	                        [seq] int
                        )
                        INSERT INTO @TmpWay
                        SELECT w.[route_id]
                              ,w.[stop_id]
                              ,c.stop_name_th
                              ,w.[seq]
                          FROM [Google_Transit].[dbo].[way] w   
                          LEFT JOIN [Google_Transit].[dbo].[stop] c
	                        ON w.stop_id = c.stop_id
                          WHERE [route_id] in (SELECT [route_id]     
					                          FROM [Google_Transit].[dbo].[routes]
					                          WHERE [route_code] = @LINE 
					                          AND [is_inbound] = @IS_INBOUND)
                        ORDER by seq

                        SELECT A.[route_id]
                              ,A.[stop_id]
                              ,A.[stop_name_th]
                              ,B._div
                              ,B.DIV
                              ,B._time
                              ,B.TIME
                              ,A.[seq]
                        FROM ( SELECT A1.route_id,
			                        A1.stop_id,
			                        A1.stop_name_th,
			                        A2.stop_id as stop_id_2,
			                        A2.stop_name_th as stop_name_th_2,
			                        A1.seq
		                        FROM @TmpWay A1
		                        LEFT JOIN @TmpWay A2
			                        ON A1.seq+1 = (A2.seq)
		                        WHERE A2.stop_id is not null
                        ) A
                        LEFT JOIN [Google_Transit].[dbo].BUSSTOP_TIME_REVERT B
	                        ON A.stop_id = B.stop_id 
	                        AND A.stop_id_2 = B.stop_id_2
	                        AND B.BOUND = @INBOUND_TXT";
                #endregion
                System.Data.DataTable DT_OUTBOUND = Retreive(Sql);
                #region OUTBOUND
                if (OpenExcel(System.Windows.Forms.Application.StartupPath + "/config/Template.xls"))
                {
                    SheetData.get_Range("B3", missing).Value2 = line_code;
                    SheetData.get_Range("B4", missing).Value2 = inbound_detail;
                    SheetData.get_Range("F5", missing).Value2 = "" + DT_INBOUND.Rows.Count;
                    int StartIndex = 12;
                    for (int j = 0; j < DT_OUTBOUND.Rows.Count; j++)
                    {
                        string seq = "" + DT_OUTBOUND.Rows[j]["seq"];
                        string stop_Name_th = "" + DT_OUTBOUND.Rows[j]["stop_Name_th"];
                        string _div = "" + DT_OUTBOUND.Rows[j]["_div"];
                        string DIV = "" + DT_OUTBOUND.Rows[j]["DIV"];
                        string _time = "" + DT_OUTBOUND.Rows[j]["_time"];
                        string TIME = "" + DT_OUTBOUND.Rows[j]["TIME"];

                        DIV = DIV.Replace("km", "");
                        SheetData.get_Range("A" + (StartIndex + j), missing).Value2 = seq;
                        SheetData.get_Range("B" + (StartIndex + j), missing).Value2 = stop_Name_th;
                        SheetData.get_Range("D" + (StartIndex + j), missing).Value2 = DIV.Trim();
                        SheetData.get_Range("F" + (StartIndex + j), missing).Value2 = _time;
                    }
                    SaveExcel(System.Windows.Forms.Application.StartupPath + "/output_excel/" + line_code + "(OUTBOUND).xls");
                }
                CloseExcel();
                #endregion
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                fileInfo = dirInfo.GetFiles("*.kml");

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(RunLoadKml);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;
            }
        }     
        private void button4_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                alError = new ArrayList();
                string FilePath = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dirInfo = new DirectoryInfo(FilePath);
                fileInfo = dirInfo.GetFiles("*.xls");
                if (fileInfo.Length == 0)
                {
                    fileInfo = dirInfo.GetFiles("*.xlsx");
                }
                isINBOUND = false;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(RunLoadExcel);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;

                button4.Text = "3. OUTBOUND นำเข้าข้อมูลจาก Template (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
            }              
        }
        private void button9_Click(object sender, EventArgs e)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(RunGen_GTFS_ALL);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
            frmWait = new frmProcess();
            bw.RunWorkerAsync();
            frmWait.ShowDialog();
            frmWait = null;            
        }  
        private void button10_Click(object sender, EventArgs e)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(RunGen_GTFS);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
            frmWait = new frmProcess();
            bw.RunWorkerAsync();
            frmWait.ShowDialog();
            frmWait = null;
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string Sql = @"SELECT [line]
                                  ,[detail]
                              FROM [Google_Transit].[dbo].[A009]
                              GROUP BY [line]
                                  ,[detail]";
            System.Data.DataTable DT_Line = Retreive(Sql);

            for (int i = 0; i < DT_Line.Rows.Count; i++)
            {
                string line_code = ""+DT_Line.Rows[i]["line"];
                string detail = ""+DT_Line.Rows[i]["detail"];

                Sql = @"SELECT [seq]
                              ,[stop_name_th]
                              ,[stop_name_en]
                              ,[start_time]
                              ,[end_time]
                          FROM [Google_Transit].[dbo].[A009]
                          WHERE line = '"+line_code+"'  ORDER BY seq";
                System.Data.DataTable dt_station = Retreive(Sql);

                if (OpenExcel(System.Windows.Forms.Application.StartupPath + "/config/Template.xls"))
                {
                    SheetData.get_Range("B2", missing).Value2 = "A009";
                    SheetData.get_Range("B3", missing).Value2 = line_code;
                    SheetData.get_Range("B4", missing).Value2 = detail;
                    SheetData.get_Range("G5", missing).Value2 = "" + dt_station.Rows.Count;
                    int StartIndex = 12;
                    for (int j = 0; j < dt_station.Rows.Count; j++)
                    {
                        string seq = "" + dt_station.Rows[j]["seq"];
                        string stop_Name_th = "" + dt_station.Rows[j]["stop_Name_th"];
                        string start_time = "" + dt_station.Rows[j]["start_time"];
                        string end_time = "" + dt_station.Rows[j]["end_time"];

                        SheetData.get_Range("A" + (StartIndex + j), missing).Value2 = seq;
                        SheetData.get_Range("B" + (StartIndex + j), missing).Value2 = stop_Name_th;
                        SheetData.get_Range("I" + (StartIndex + j), missing).Value2 = start_time;
                        SheetData.get_Range("J" + (StartIndex + j), missing).Value2 = end_time;

                    }
                    SaveExcel(System.Windows.Forms.Application.StartupPath + "/output_excel/" + line_code + ".xls");
                }
                CloseExcel();
            }
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            //cbxAgency
            System.Data.DataTable DTAgency = new System.Data.DataTable();
            DTAgency.Columns.Add("Code");
            DTAgency.Columns.Add("Text");

            DTAgency.Rows.Add(AddRow(DTAgency, "A001", "A001 BTS"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A002", "A002 MRT"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A003", "A003 ขสมก"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A004", "A004 รถตู้"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A005", "A005 เรือด่วนเจ้าพระยา"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A006", "A006 SRT"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A007", "A007 BRT"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A008", "A008 เรือโดยสารคลองแสนแสบ"));
            DTAgency.Rows.Add(AddRow(DTAgency, "A009", "A009 รถไฟ"));

            cbxAgency.DataSource = DTAgency.Copy();
            cbxAgency.DisplayMember = "Text";
            cbxAgency.ValueMember = "Code";

            cbxAgency2.DataSource = DTAgency.Copy();
            cbxAgency2.DisplayMember = "Text";
            cbxAgency2.ValueMember = "Code";

            txbFolderKML.Text = System.Windows.Forms.Application.StartupPath + @"\output_kml";
            txbFolderGTFS.Text = System.Windows.Forms.Application.StartupPath + @"\output";
        }
        #endregion


        private void fnGenKMLAll()
        {
            string Sql = @"SELECT [name]
                                  ,[lon]
                                  ,[lat] 
                              FROM [Google_Transit].[dbo].[BUSSTOPALL$]";
            System.Data.DataTable DT = Retreive(Sql);
            string strData = "<Folder><name>1</name><open>1</open><Folder><name>ALL</name>";
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string strTmp = @"<Placemark>
				                        <name>@name</name>
				                        <open>1</open>
				                        <LookAt>
					                        <longitude>@longitude</longitude>
					                        <latitude>@latitude</latitude>
					                        <altitude>0</altitude>
					                        <heading>0.006033119647067864</heading>
					                        <tilt>0</tilt>
					                        <range>200</range>
					                        <altitudeMode>relativeToGround</altitudeMode>
				                        </LookAt>
				                        <styleUrl>#m_ylw-pushpin10</styleUrl>
				                        <Point>
					                        <coordinates>@longitude,@latitude,0</coordinates>
				                        </Point>
			                        </Placemark>";
                strTmp = strTmp.Replace("@name", "" + DT.Rows[i]["name"]);
                strTmp = strTmp.Replace("@longitude", "" + DT.Rows[i]["lon"]);
                strTmp = strTmp.Replace("@latitude", "" + DT.Rows[i]["lat"]);
                strData += strTmp;      
            }
            strData += "</Folder></Folder>";

            string outputXML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
            outputXML += "<kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:kml=\"http://www.opengis.net/kml/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\">";
            outputXML += "<Document>";
            outputXML += "<name>BUSSTOP_ALL.kml</name>";
            outputXML += File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/config/style_kml.txt");
            outputXML += strData;
            outputXML += "</Document>";
            outputXML += "</kml>";


            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output_kml/BUSSTOP_ALL.kml", outputXML);
        }
        


        #region Function Process
        frmProcess frmWait;
        string MsgOutput;
        FileInfo[] fileInfo;
        string FilePathTmp = "";
        bool isINBOUND;
        private void RunLoadKml(object sender, DoWorkEventArgs e)
        {
            string AgencyID = "";
            this.Invoke(
               (MethodInvoker)delegate()
               {
                   frmWait.labCount.Text = "0/" + fileInfo.Length;
                   frmWait.pgbMain.Maximum = fileInfo.Length;
                   frmWait.pgbMain.Step = 1;
                   AgencyID = "" + cbxAgency.SelectedValue;
               }
           );


            for (int i = 0; i < fileInfo.Length; i++)
            {
                fnLoad_KML(fileInfo[i].FullName, AgencyID, "0");
                this.Invoke(
                      (MethodInvoker)delegate()
                      {
                          frmWait.labCount.Text = (i + 1) + "/" + fileInfo.Length;
                          frmWait.pgbMain.PerformStep();
                      }
                  );
            }


            string Sql = @"UPDATE [Google_Transit].[dbo].[shapes] SET line_id =  REPLACE([line_id],'.kml','')
                            UPDATE [Google_Transit].[dbo].[shapes] SET line_id =  REPLACE([line_id],'รถร่วม','ร')
                            UPDATE [Google_Transit].[dbo].[shapes] SET line_id =  REPLACE([line_id],'เสริม','ส')
                            UPDATE [Google_Transit].[dbo].[shapes] SET line_id =  REPLACE([line_id],'ทางด่วน','ทด')
                            UPDATE [Google_Transit].[dbo].[shapes] SET line_id =  REPLACE([line_id],'ส1','ส')";

            string ErrorMSG = "";
            if (!Execute(Sql, out ErrorMSG))
                alError.Add("UPDATE shapes => Error:" + ErrorMSG);

            MsgOutput = "";
            for (int i = 0; i < alError.Count; i++)
            {
                MsgOutput += alError[i] + Environment.NewLine;
            }

            if (MsgOutput == "")
            {
                MsgOutput = "นำเข้าข้อมูล KML เรียบร้อยแล้ว";
            }
        }
        private void RunComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            if (frmWait != null)
            {
                frmWait.Close();
                frmWait = null;
            }

            Random RD = new Random();
            string FileName = DateTime.Now.ToShortDateString();
            FileName = FileName.Replace("/", "-") + "_" + RD.Next(1000, 9999).ToString() + ".txt";
            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/log/" + FileName, MsgOutput);
            MessageBox.Show(MsgOutput, "ผลการทำงาน", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RunLoadExcel(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
              (MethodInvoker)delegate()
              {
                  frmWait.labCount.Text = "0/" + fileInfo.Length;
                  frmWait.pgbMain.Maximum = fileInfo.Length;
                  frmWait.pgbMain.Step = 1;
              }
          );


            for (int i = 0; i < fileInfo.Length; i++)
            {
                LoadTemplate(fileInfo[i].FullName, isINBOUND);
                this.Invoke(
                     (MethodInvoker)delegate()
                     {
                         frmWait.labCount.Text = (i + 1) + "/" + fileInfo.Length;
                         frmWait.pgbMain.PerformStep();
                     }
                 );
            }

            MsgOutput = "";
            for (int i = 0; i < alError.Count; i++)
            {
                MsgOutput += alError[i] + Environment.NewLine;
            }

            if (MsgOutput == "")
            {
                MsgOutput = "นำเข้าข้อมูล KML เรียบร้อยแล้ว";
            }
        }
       


        private void RunGen_GTFS_ALL(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
              (MethodInvoker)delegate()
              {
                  frmWait.labCount.Text = "";                 
                  frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
              }
          );
            genGTFS_All();
        }
        private void RunGen_GTFS(object sender, DoWorkEventArgs e)
        {
            string AgencyID = "";
            this.Invoke(
               (MethodInvoker)delegate()
               {
                   frmWait.labCount.Text = "";
                   AgencyID = "" + cbxAgency2.SelectedValue;
                   frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
               }
           );
            genGTFS_By_Agency(AgencyID);
        }


        
        #endregion

        private void button12_Click(object sender, EventArgs e)
        {
            fnGenKMLAll();
        }

       
        void RunFindFeeds(object sender, DoWorkEventArgs e)
        {
            string ErrorMSG = "";
            string Sql = @"SELECT [line_id]
                          FROM [Google_Transit].[dbo].[line]
                          GROUP BY [line_id]";
            System.Data.DataTable DT_Line = Retreive(Sql);

            this.Invoke(
               (MethodInvoker)delegate()
               {
                   frmWait.labCount.Text = "0/" + DT_Line.Rows.Count;
                   frmWait.pgbMain.Maximum = DT_Line.Rows.Count;
                   frmWait.pgbMain.Step = 1;
               }
            );

            Random RD = new Random();
            string round_no = DateTime.Now.ToString("ddMMyyyy") + "_" + RD.Next(1000, 9999);

            for (int i = 0; i < DT_Line.Rows.Count; i++)
            {
                string line_id = "" + DT_Line.Rows[i]["line_id"];
                Sql = @"SELECT  A.round_no
		                        ,A.stop_id 
		                        ,A.stop_sequence	
		                        ,B.stop_id as [stop_id_2]
		                        ,b.stop_sequence as [stop_sequence_2]
		                        ,A.stop_lat
		                        ,A.stop_lon
		                        ,B.stop_lat as [stop_lat_2]
		                        ,B.stop_lon as [stop_lon_2]
		                        ,A.RowId
                        FROM (
	                        SELECT A.line_id,A.is_inbound,A.round_no,A.service_id,A.stop_sequence,A.stop_id
		                          ,B.stop_name_th,B.stop_lat,B.stop_lon,A.RowId,A.is_feeds
	                        FROM [stop_times] A
	                        LEFT JOIN [stop] B
		                        ON A.stop_id = B.stop_id
                        ) AS A 
                        LEFT JOIN (
	                        SELECT A.line_id,A.is_inbound,A.round_no,A.service_id,A.stop_sequence,A.stop_id
		                          ,B.stop_name_th,B.stop_lat,B.stop_lon
	                        FROM [stop_times] A
	                        LEFT JOIN [stop] B
		                        ON A.stop_id = B.stop_id
                        ) AS B 
	                        ON A.line_id = B.line_id 
	                        AND A.is_inbound = B.is_inbound 
	                        AND A.round_no = B.round_no 
	                        AND A.service_id = B.service_id 
	                        AND B.stop_sequence = A.stop_sequence + 1 
                        WHERE A.line_id = '@line_id'
                        AND B.stop_sequence is not null
                        AND isnull(A.is_feeds,0) <> 1
                        ORDER BY A.is_inbound,A.round_no,A.service_id,A.stop_sequence";
                Sql = Sql.Replace("@line_id", line_id);
                System.Data.DataTable DT_TMP = Retreive(Sql);

                for (int j = 0; j < DT_TMP.Rows.Count; j++)
                {
                    string stop_lat = "" + DT_TMP.Rows[j]["stop_lat"];
                    string stop_lon = "" + DT_TMP.Rows[j]["stop_lon"];

                    string stop_lat_2 = "" + DT_TMP.Rows[j]["stop_lat_2"];
                    string stop_lon_2 = "" + DT_TMP.Rows[j]["stop_lon_2"];

                    string RowId = "" + DT_TMP.Rows[j]["RowId"];
                    if (!fund_point(stop_lat, stop_lon, stop_lat_2, stop_lon_2))
                    {
                        string start_point = stop_lat + "," + stop_lon;
                        string end_point = stop_lat_2 + "," + stop_lon_2;

                        DataSet ds = new DataSet();
                        ds.ReadXml("http://maps.google.com/maps/api/directions/xml?origin=" + start_point + "&destination=" + end_point + "&sensor=false");

                        string status = "" + ds.Tables["DirectionsResponse"].Rows[0]["status"];
                        Sql = @"INSERT INTO [Google_Transit].[dbo].[Feeds_Fail]
                                               ([stop_lat]
                                               ,[stop_lon]
                                               ,[stop_lat_2]
                                               ,[stop_lon_2]
                                               ,[msg]
                                               ,[create_date]
                                               ,[round_no])
                                         VALUES
                                               ('@stop_lat'
                                               ,'@stop_lon'
                                               ,'@stop_lat_2'
                                               ,'@stop_lon_2'
                                               ,'@msg'
                                               ,GETDATE()
                                               ,'@round_no')";
                        Sql = Sql.Replace("@stop_lat_2", stop_lat_2);
                        Sql = Sql.Replace("@stop_lon_2", stop_lon_2);
                        Sql = Sql.Replace("@stop_lat", stop_lat);
                        Sql = Sql.Replace("@stop_lon", stop_lon);
                        Sql = Sql.Replace("@msg", status);
                        Sql = Sql.Replace("@round_no", round_no);
                        
                        if (!Execute(Sql, out ErrorMSG))
                            alError.Add("INSERT Feeds_Fail => Error:" + ErrorMSG);


                        if (status == "OVER_QUERY_LIMIT")
                        {
                            System.Threading.Thread.Sleep(10000);
                        }
                       

                        try
                        {
                            //หาระยะทาง

                            DataRow[] dr = ds.Tables["distance"].Select("step_Id is null");
                           
                            string text_distance = "";
                            string out_distance = "";
                            try {
                                out_distance = "" + dr[0]["value"];
                                text_distance = "" + dr[0]["text"];
                            }
                            catch {  }

                            
                            //หาเวลาที่ใช้
                            System.Data.DataTable dt = ds.Tables["duration"];
                            int out_time = 99999999;
                            string text_time = "";
                            for (int K = 0; K < dt.Rows.Count; K++)
                            {
                                int time = strToInt(dt.Rows[K]["value"].ToString());
                                if (time < out_time)
                                {
                                    out_time = time;
                                    text_time = "" + dt.Rows[K]["text"];
                                }
                            }
                            //หารายละเอียดเส้นทาง
                            string route_txt = "" + ds.Tables["route"].Rows[0]["summary"];
                            string start_address = "" + ds.Tables["leg"].Rows[0]["start_address"];
                            string end_address = "" + ds.Tables["leg"].Rows[0]["end_address"];

                            Sql = @"INSERT INTO [Google_Transit].[dbo].[Feeds]
                                           ([route_txt] ,[start_address],[end_address]
                                           ,[stop_lat],[stop_lon],[stop_lat_2] ,[stop_lon_2]
                                           ,[distance] ,[distance_txt],[time] ,[time_txt])
                                     VALUES
                                           ('@route_txt'
                                           ,'@start_address'
                                           ,'@end_address'
                                           ,'@stop_lat'
                                           ,'@stop_lon'
                                           ,'@stop_lat_2'
                                           ,'@stop_lon_2'
                                           ,'@distance'
                                           ,'@distance_txt'
                                           ,'@time'
                                           ,'@time_txt')";
                            Sql = Sql.Replace("@route_txt", route_txt);
                            Sql = Sql.Replace("@start_address", start_address);
                            Sql = Sql.Replace("@end_address", end_address);
                            Sql = Sql.Replace("@stop_lat_2", stop_lat_2);
                            Sql = Sql.Replace("@stop_lon_2", stop_lon_2);
                            Sql = Sql.Replace("@stop_lat", stop_lat);
                            Sql = Sql.Replace("@stop_lon", stop_lon);
                            Sql = Sql.Replace("@distance_txt", text_distance);
                            Sql = Sql.Replace("@distance", "" + out_distance);
                            Sql = Sql.Replace("@time_txt", text_time);
                            Sql = Sql.Replace("@time", "" + out_time);

                            ErrorMSG = "";
                            if (!Execute(Sql, out ErrorMSG))
                                alError.Add("INSERT Feeds => route_txt:" + route_txt + ",start_point:" + start_point + ",end_point:" + end_point + ",Error:" + ErrorMSG);

                            
                            Sql = "UPDATE stop_times SET is_feeds = 1 WHERE RowId = " + RowId;                           
                            if (!Execute(Sql, out ErrorMSG))
                                alError.Add("UPDATE Feeds => Error:" + ErrorMSG);

                            //System.Threading.Thread.Sleep(1000);

                        }
                        catch (Exception ex)
                        {
                            alError.Add("GOOGLE Feeds => Error:" + ex.Message);
                        }
                    }
                    else
                    {
                        alError.Add("Found int Feeds");
                        Sql = "UPDATE stop_times SET is_feeds = 1 WHERE RowId = " + RowId;
                        ErrorMSG = "";
                        if (!Execute(Sql, out ErrorMSG))
                            alError.Add("UPDATE Feeds => Error:" + ErrorMSG);
                    }
                }

                this.Invoke(
                     (MethodInvoker)delegate()
                     {
                         frmWait.labCount.Text = (i + 1) + "/" + DT_Line.Rows.Count;
                         frmWait.pgbMain.PerformStep();
                     }
                 );
            }
        }
        bool fund_point(string stop_lat, string stop_lon, string stop_lat_2, string stop_lon_2)
        {
            string Sql = @"SELECT COUNT(*) as [COUNT]
                          FROM [Google_Transit].[dbo].[Feeds]
                          WHERE [stop_lat] = '@stop_lat'
                          AND [stop_lon] = '@stop_lon'
                          AND [stop_lat_2] = '@stop_lat_2'
                          AND [stop_lon_2] = '@stop_lon_2'";
            Sql = Sql.Replace("@stop_lat_2", stop_lat_2);
            Sql = Sql.Replace("@stop_lon_2", stop_lon_2);
            Sql = Sql.Replace("@stop_lat", stop_lat);
            Sql = Sql.Replace("@stop_lon", stop_lon);
            System.Data.DataTable DT_TMP = Retreive(Sql);
            if ("" + DT_TMP.Rows[0][0] == "0")
                return false;
            else
                return true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("คุณแน่ใจหรือไม่ที่จะทำการลบข้อมูลนี้ ?", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                deleteData();
                button14.Text = "1. ล้างข้อมูลก่อนการประมวลผล (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
            }            
        }

        private void deleteData()
        {
            string Sql = @"exec [dbo].[sp_Delete_All] 'DeleteData'";
            string ErrorLog = "";
            if (Execute(Sql, out ErrorLog))
                MessageBox.Show("ลบข้อมูลเรียบร้อยแล้ว", "ผลการทำงาน", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(ErrorLog, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);

            
        }
        private void updatePoint()
        {
            string Sql = @"exec [dbo].[sp_update_lat_lon_from_busstopall]";
            string ErrorLog = "";
            if (Execute(Sql, out ErrorLog))
                MessageBox.Show("อัพเดทพิกัดเรียบร้อยแล้ว", "ผลการทำงาน", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(ErrorLog, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void updateTime(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
                  (MethodInvoker)delegate()
                  {
                      frmWait.labCount.Text = "";
                      frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
                  }
              );
            string Sql = @"exec [dbo].[spl_gen_time]";
            string ErrorLog = "";
            if (Execute(Sql, out ErrorLog))
                MsgOutput = "อัพเดทพิกัดเรียบร้อยแล้ว";
            else
                MsgOutput = ErrorLog;

        }
      
        private void button15_Click(object sender, EventArgs e)
        {
            updatePoint();

            button15.Text = "4. อัพเดทพิกัดจาก Busstop All (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(RunFindFeeds);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
            frmWait = new frmProcess();
            bw.RunWorkerAsync();
            frmWait.ShowDialog();
            frmWait = null;
            alError = new ArrayList();

            button13.Text = "5.1 หาระยะทางเพื่อเตรียมประมวลผลอัพเดทเวลา (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
        }

        private void button16_Click(object sender, EventArgs e)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(updateTime);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
            frmWait = new frmProcess();
            bw.RunWorkerAsync();
            frmWait.ShowDialog();
            frmWait = null;

            button16.Text = "5.2 ประมวลผลเวลาโดยใช้สูตร (" + DateTime.Now.ToLongDateString() + " เวลา " + DateTime.Now.ToLongTimeString() + ")";
        }

        private void button17_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                alError = new ArrayList();
                FilePathTmp = ofd.FileName;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(loadBusstopAll);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;
            }            
        }

        void loadBusstopAll(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
                 (MethodInvoker)delegate()
                 {
                     frmWait.labCount.Text = "";
                     frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
                 }
             );

            if (OpenExcel(FilePathTmp))
            {
                string Sql = "DELETE [Google_Transit].[dbo].[BUSSTOPALL$]";
                string ErrorMSG = "";

                if (!Execute(Sql, out ErrorMSG))
                {
                    MsgOutput = "DELETE BUSSTOPALL$ => Error:" + ErrorMSG;
                }
                else
                {
                    for (int i = 2; i < SheetData.Rows.Count; i++)
                    {
                        string name = ("" + SheetData.get_Range("A" + i, missing).Value2).Trim();
                        string lon = ("" + SheetData.get_Range("B" + i, missing).Value2).Trim();
                        string lat = ("" + SheetData.get_Range("C" + i, missing).Value2).Trim();
                        string agency = ("" + SheetData.get_Range("D" + i, missing).Value2).Trim();

                        if (name == "" && lon == "" && lat == "" && agency == "")
                            break;
                        else
                        {
                            Sql = @"INSERT INTO [Google_Transit].[dbo].[BUSSTOPALL$]
                                               ([name]
                                               ,[lon]
                                               ,[lat]
                                               ,[agency])
                                         VALUES
                                               ('@name'
                                               ,'@lon'
                                               ,'@lat'
                                               ,'@agency')";
                            Sql = Sql.Replace("@name", name);
                            Sql = Sql.Replace("@lon", lon);
                            Sql = Sql.Replace("@lat", lat);
                            Sql = Sql.Replace("@agency", agency);

                            if (!Execute(Sql, out ErrorMSG))
                                alError.Add("INSERT BUSSTOPALL$ => name:" + name + ",Error:" + ErrorMSG);
                        }
                    }
                }
            }
            CloseExcel();

            MsgOutput = "นำเข้าข้อมูล Busstop all เรียบร้อยแล้ว";
        }

        private void button18_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                FilePathTmp = sfd.FileName;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(Report_BusstopAll);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;
            }            
        }

        
        private void button19_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                FilePathTmp = sfd.FileName;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(Report_Time1);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;
            }            
        }


        void Report_BusstopAll(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
                 (MethodInvoker)delegate()
                 {
                     frmWait.labCount.Text = "";
                     frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
                 }
             );
            if (OpenExcel(System.Windows.Forms.Application.StartupPath + @"/config/Report_Busstopall.xls"))
            {
                string Sql = @"SELECT [stop_id]
                                      ,[stop_name_th]
                                      ,[stop_sequence]
                                      ,[stop_detail]
                                      ,[stop_lat]
                                      ,[stop_lon]
                                  FROM [Google_Transit].[dbo].[v_check_busstopall]";
                System.Data.DataTable DT_TMP = Retreive(Sql);

                if (DT_TMP.Rows.Count == 0)
                {
                    MsgOutput = "ไม่พบรายการข้อผิดพลาดนี้";
                    CloseExcel();
                    return;
                }
                else
                {
                    int Index = 2;
                    for (int i = 0; i < DT_TMP.Rows.Count; i++)
                    {
                        SheetData.get_Range("A" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_id"];
                        SheetData.get_Range("B" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_name_th"];
                        SheetData.get_Range("C" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_sequence"];
                        SheetData.get_Range("D" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_detail"];
                        SheetData.get_Range("E" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_lat"];
                        SheetData.get_Range("F" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_lon"];
                        Index++;
                    }
                }

                SaveExcel(FilePathTmp);
                MsgOutput = "สร้างรายงานข้อผิดพลาดเรียบร้อยแล้ว กรุณาไปหาจาก Path ที่คุณบันทึก";
                //string name = ("" + SheetData.get_Range("A" + i, missing).Value2).Trim();

            }

            CloseExcel();


        }


        void Report_Time1(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
                 (MethodInvoker)delegate()
                 {
                     frmWait.labCount.Text = "";
                     frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
                 }
             );
            if (OpenExcel(System.Windows.Forms.Application.StartupPath + @"/config/Report_Time1.xls"))
            {
                string Sql = @"SELECT [line_code]
                                      ,[inbound_detail]
                                      ,[stop_id]
                                      ,[stop_sequence]
                                      ,[arrival_time]
                                      ,[departure_time]
                                      ,[FileName]
                                  FROM [Google_Transit].[dbo].[v_check_time]";
                System.Data.DataTable DT_TMP = Retreive(Sql);

                if (DT_TMP.Rows.Count == 0)
                {
                    MsgOutput = "ไม่พบรายการข้อผิดพลาดนี้";
                    CloseExcel();
                    return;
                }
                else
                {
                    int Index = 2;
                    for (int i = 0; i < DT_TMP.Rows.Count; i++)
                    {
                        SheetData.get_Range("A" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["line_code"];
                        SheetData.get_Range("B" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["inbound_detail"];
                        SheetData.get_Range("C" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_id"];
                        SheetData.get_Range("D" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_sequence"];
                        SheetData.get_Range("E" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["arrival_time"];
                        SheetData.get_Range("F" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["departure_time"];
                        SheetData.get_Range("G" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["FileName"];
                        Index++;
                    }
                }

                SaveExcel(FilePathTmp);
                MsgOutput = "สร้างรายงานข้อผิดพลาดเรียบร้อยแล้ว กรุณาไปหาจาก Path ที่คุณบันทึก";
            }
            CloseExcel();
        }

        void Report_Time2(object sender, DoWorkEventArgs e)
        {
            this.Invoke(
                 (MethodInvoker)delegate()
                 {
                     frmWait.labCount.Text = "";
                     frmWait.pgbMain.Style = ProgressBarStyle.Marquee;
                 }
             );
            if (OpenExcel(System.Windows.Forms.Application.StartupPath + @"/config/Report_Time2.xls"))
            {
                string Sql = @"SELECT [line_code]
                                      ,[inbound_detail]
                                      ,[stop_id]
                                      ,[stop_sequence]
                                      ,[arrival_time]
                                      ,[departure_time]
                                      ,[Expr1]
                                      ,[FileName]
                                  FROM [Google_Transit].[dbo].[v_check_time2]";
                System.Data.DataTable DT_TMP = Retreive(Sql);

                if (DT_TMP.Rows.Count == 0)
                {
                    MsgOutput = "ไม่พบรายการข้อผิดพลาดนี้";
                    CloseExcel();
                    return;
                }
                else
                {
                    int Index = 2;
                    for (int i = 0; i < DT_TMP.Rows.Count; i++)
                    {
                        SheetData.get_Range("A" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["line_code"];
                        SheetData.get_Range("B" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["inbound_detail"];
                        SheetData.get_Range("C" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_id"];
                        SheetData.get_Range("D" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["stop_sequence"];
                        SheetData.get_Range("E" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["arrival_time"];
                        SheetData.get_Range("F" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["departure_time"];
                        SheetData.get_Range("G" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["Expr1"];
                        SheetData.get_Range("H" + Index, missing).Value2 = "" + DT_TMP.Rows[i]["FileName"];
                        Index++;
                    }
                }

                SaveExcel(FilePathTmp);
                MsgOutput = "สร้างรายงานข้อผิดพลาดเรียบร้อยแล้ว กรุณาไปหาจาก Path ที่คุณบันทึก";
            }
            CloseExcel();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                FilePathTmp = sfd.FileName;

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(Report_Time2);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunComplete);
                frmWait = new frmProcess();
                bw.RunWorkerAsync();
                frmWait.ShowDialog();
                frmWait = null;
            }            
        }

        private void button21_Click(object sender, EventArgs e)
        {
            string Sql = @"exec sp_clean_time";
            string ErrorLog = "";
            if (Execute(Sql, out ErrorLog))
                MessageBox.Show("จัดการข้อมูลเวลาเรียบร้อยแล้ว", "ผลการทำงาน", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(ErrorLog, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private double distance(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            double theta = lon1 - lon2;
            double dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta));
            dist = Math.Acos(dist);
            dist = rad2deg(dist);
            dist = dist * 60 * 1.1515;
            if (unit == 'K')
            {
                dist = dist * 1.609344;
            }
            else if (unit == 'N')
            {
                dist = dist * 0.8684;
            }
            return (dist);
        }
        private double deg2rad(double deg)
        {
            return (deg * Math.PI / 180.0);
        }
        private double rad2deg(double rad)
        {
            return (rad / Math.PI * 180.0);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            string Sql = @"SELECT RowId,stop_lat,stop_lon,stop_lat_2,stop_lon_2,distance
                              FROM [Google_Transit].[dbo].[Feeds]                              
                              ORDER BY RowId";
            System.Data.DataTable dt = Retreive(Sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string RowId = "" + dt.Rows[i]["RowId"];
                string stop_lat = "" + dt.Rows[i]["stop_lat"];
                string stop_lon = "" + dt.Rows[i]["stop_lon"];
                string stop_lat_2 = "" + dt.Rows[i]["stop_lat_2"];
                string stop_lon_2 = "" + dt.Rows[i]["stop_lon_2"];

                double distance_cal = distance(strToDouble(stop_lat), strToDouble(stop_lon), strToDouble(stop_lat_2), strToDouble(stop_lon_2),'K')*1000;

                Sql = @"UPDATE [Google_Transit].[dbo].[Feeds]
                        SET distance_cal = '" + distance_cal + @"'
                        WHERE RowId = " + RowId;
                string ErrorMsg = "";
                Execute(Sql, out ErrorMsg);
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            string Sql = @"SELECT [line_id] ,[line_code]   
                                  ,[inbound_detail]
                                  ,[outbound_detail]                                      
                              FROM [Google_Transit].[dbo].[line]
                            WHERE agency_id = 'A003'
                              GROUP BY [line_id] ,[line_code] 
                                  ,[inbound_detail]
                                  ,[outbound_detail]";
            System.Data.DataTable DT = Retreive(Sql);

          
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string line = "" + DT.Rows[i]["line_id"];
                string detail = "" + DT.Rows[i]["inbound_detail"];
                string line_code = "" + DT.Rows[i]["line_code"];

                #region Inbound
                Sql = @"SELECT A.[line_id]
                              ,A.[is_inbound]    
                              ,A.[stop_id]    
                              ,A.[stop_sequence] 
                              ,B.[stop_name_th]
                              ,B.[stop_lat]
                              ,B.[stop_lon]   
                          FROM [Google_Transit].[dbo].[stop_times] A
                          LEFT JOIN [Google_Transit].[dbo].[stop] B
	                        ON A.stop_id = B.stop_id
                          WHERE A.is_inbound = 1
                          and A.line_id = '@line'
                          ORDER BY A.stop_sequence";
                Sql = Sql.Replace("@line", line);
                System.Data.DataTable dtTemp = Retreive(Sql);

                string strData = "<Folder><name>" + line_code + "</name><open>1</open><Folder><name>INBOUND</name>";
                for (int j = 0; j < dtTemp.Rows.Count; j++)
                {
                    string strTmp = @"<Placemark>
				                        <name>[ INBOUND ]@name</name>
				                        <open>1</open>
				                        <LookAt>
					                        <longitude>@longitude</longitude>
					                        <latitude>@latitude</latitude>
					                        <altitude>0</altitude>
					                        <heading>0.006033119647067864</heading>
					                        <tilt>0</tilt>
					                        <range>200</range>
					                        <altitudeMode>relativeToGround</altitudeMode>
				                        </LookAt>
				                        <styleUrl>#m_ylw-pushpin10</styleUrl>
				                        <Point>
					                        <coordinates>@longitude,@latitude,0</coordinates>
				                        </Point>
			                        </Placemark>";
                    strTmp = strTmp.Replace("@name", "" + dtTemp.Rows[j]["stop_name_th"]);
                    strTmp = strTmp.Replace("@longitude", "" + dtTemp.Rows[j]["stop_lon"]);
                    strTmp = strTmp.Replace("@latitude", "" + dtTemp.Rows[j]["stop_lat"]);

                    strData += strTmp;
                }

                Sql = @"SELECT [shape_pt_lon]+','+[shape_pt_lat]+',0' as [point] 
                          FROM [Google_Transit].[dbo].[shapes]
                          WHERE is_inbound = 1
                          AND line_id = '"+line_code+@"'
                          ORDER BY stop_sequence";
                dtTemp = Retreive(Sql);
                string SPOINT = "";
                for (int j = 0; j < dtTemp.Rows.Count; j++)
                {
                    if (SPOINT == "") SPOINT = "" + dtTemp.Rows[j]["point"];
                    else SPOINT += " " + dtTemp.Rows[j]["point"];
                }
                strData += @"<Placemark>
				                 <name>ขาเข้า 1</name>
				                 <description>"+detail+@"</description>
				                 <styleUrl>#stylemap_id10</styleUrl>
				                 <LineString>
					                 <tessellate>1</tessellate>
					                 <coordinates>"+SPOINT+@"</coordinates>
				                 </LineString>
			                 </Placemark>";

                strData += "</Folder>";

                #endregion

                #region OUTBOUND
               
                Sql = @"SELECT A.[line_id]
                              ,A.[is_inbound]    
                              ,A.[stop_id]    
                              ,A.[stop_sequence] 
                              ,B.[stop_name_th]
                              ,B.[stop_lat]
                              ,B.[stop_lon]   
                          FROM [Google_Transit].[dbo].[stop_times] A
                          LEFT JOIN [Google_Transit].[dbo].[stop] B
	                        ON A.stop_id = B.stop_id
                          WHERE A.is_inbound = 0
                          and A.line_id = '@line'
                          ORDER BY A.stop_sequence";
                Sql = Sql.Replace("@line", line);
                dtTemp = Retreive(Sql);

                strData += "<Folder><name>OUTBOUND</name>";
                for (int j = 0; j < dtTemp.Rows.Count; j++)
                {
                    string strTmp = @"<Placemark>
				                        <name>[ OUTBOUND ]@name</name>
				                        <open>1</open>
				                        <LookAt>
					                        <longitude>@longitude</longitude>
					                        <latitude>@latitude</latitude>
					                        <altitude>0</altitude>
					                        <heading>0.006033119647067864</heading>
					                        <tilt>0</tilt>
					                        <range>200</range>
					                        <altitudeMode>relativeToGround</altitudeMode>
				                        </LookAt>
				                        <styleUrl>#m_ylw-pushpin10</styleUrl>
				                        <Point>
					                        <coordinates>@longitude,@latitude,0</coordinates>
				                        </Point>
			                        </Placemark>";
                    strTmp = strTmp.Replace("@name", "" + dtTemp.Rows[j]["stop_name_th"]);
                    strTmp = strTmp.Replace("@longitude", "" + dtTemp.Rows[j]["stop_lon"]);
                    strTmp = strTmp.Replace("@latitude", "" + dtTemp.Rows[j]["stop_lat"]);

                    strData += strTmp;
                }

                Sql = @"SELECT [shape_pt_lon]+','+[shape_pt_lat]+',0' as [point] 
                          FROM [Google_Transit].[dbo].[shapes]
                          WHERE is_inbound = 0
                          AND line_id = '" + line_code + @"'
                          ORDER BY stop_sequence";
                dtTemp = Retreive(Sql);
                SPOINT = "";
                for (int j = 0; j < dtTemp.Rows.Count; j++)
                {
                    if (SPOINT == "") SPOINT = "" + dtTemp.Rows[j]["point"];
                    else SPOINT += " " + dtTemp.Rows[j]["point"];
                }
                strData += @"<Placemark>
				                 <name>ขาออก 1</name>
				                 <description>" + detail + @"</description>
				                 <styleUrl>#stylemap_id10</styleUrl>
				                 <LineString>
					                 <tessellate>1</tessellate>
					                 <coordinates>" + SPOINT + @"</coordinates>
				                 </LineString>
			                 </Placemark>";
                strData += "</Folder>";

                #endregion

                strData += "</Folder>";
                string outputXML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
                outputXML += "<kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:kml=\"http://www.opengis.net/kml/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\">";
                outputXML += "<Document>";
                outputXML += "<name>" + line_code + ".kml</name>";
                outputXML += File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/config/style_kml.txt");
                outputXML += strData;
                outputXML += "</Document>";
                outputXML += "</kml>";


                File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/output_kml/" + line_code + " [" + detail + "].kml", outputXML);
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            string[] line = System.IO.File.ReadAllLines("gtfs/translations.txt");
            alError = new ArrayList();
            for (int i = 0; i < line.Length; i++)
            {
                string[] SP = line[i].Split(new char[] { ',' });
                string Sql = @"INSERT INTO [dbo].[translations]
                                   ([trans_id]
                                   ,[lang]
                                   ,[translation])
                             VALUES
                                   ('@trans_id'
                                   ,'@lang'
                                   ,'@translation')";
                Sql = Sql.Replace("@trans_id", SP[0].Replace("'","''"));
                Sql = Sql.Replace("@lang", SP[1].Replace("'", "''"));
                Sql = Sql.Replace("@translation", SP[2].Replace("'", "''"));


                string ErrorMSG = "";
                if (!Execute(Sql, out ErrorMSG))
                    alError.Add("INSERT translations => Error:" + ErrorMSG);
            }
        }

        /*
        AgencyID	AgencyName	agency_phone	agency_url
        A001      	BTS	026177300	https://www.bts.co.th/customer/th/main.aspx
        A002      	MRT	023542000	http://www.bemplc.co.th
        A003      	ขสมก	022460973	http://www.bmta.co.th
        A004      	รถตู้		http://www.where-in-thailand.com/transit
        A005      	เรือด่วนเจ้าพระยา	024458888	http://www.chaophrayaexpressboat.com
        A006      	SRT	1690	http://www.srtet.co.th
        A007      	BRT	026177300	http://www.where-in-thailand.com/transit
        A008      	เรือโดยสารคลองแสนแสบ	023748990	http://www.where-in-thailand.com/transit
        A009      	รถไฟ	1690	http://www.railway.co.th
        
        route_type
        0: รถราง, รางรถรางเบา ใด ๆ ระบบรถไฟหรือถนนระดับแสงภายในพื้นที่กรุงเทพมหานครและปริมณฑล
        1: รถไฟใต้ดินรถไฟใต้ดิน ใด ๆ ที่ระบบรถไฟใต้ดินในพื้นที่กรุงเทพมหานครและปริมณฑล
        2: รถไฟ ใช้สำหรับการระหว่างเมืองหรือทางไกลเดินทาง
        3: รถประจำทาง ใช้สำหรับเส้นทางรถเมล์ในระยะสั้นและระยะยาว
        4: เรือเฟอร์รี่ ใช้สำหรับในระยะสั้นและระยะยาวบริการเรือ
        5: รถสายเคเบิ้ล ใช้สำหรับสายรถยนต์ระดับถนนสายที่วิ่งอยู่ใต้รถ
        6: กอนโดลา, เคเบิ้ลที่ถูกระงับรถ โดยทั่วไปจะใช้สำหรับรถยนต์สายอากาศที่รถถูกระงับจากสายเคเบิล
        7: Funicular ระบบรถไฟใด ๆ ที่ออกแบบมาสำหรับการเอียงลาดชัน
         */
    }
}
