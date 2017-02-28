using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelFormatter
{
    public class FormatterHelper
    {
        public static List<KeyValuePair<string, string>> FormatterType => new List<KeyValuePair<string, string>>()
        {
            new KeyValuePair<string, string>("A", "1. 将客户要求中的联合要求拆分为多条"),
            new KeyValuePair<string, string>("B", "2. 将实体表中包含多个实体的数据拆分为多条"),
            new KeyValuePair<string, string>("C", "3. 从卷号中拆分出案件类型代码"),
            new KeyValuePair<string, string>("D", "4. 将多条同一案件的发明人信息合并为一条")
        };

        public static void LoadExcel(string sFilePath, int nIndex, ref List<string> listSheetsName, ref DataTable dtExcelData)
        {
            var sConnectionString = $"Provider=Microsoft.Ace.OleDb.12.0;data source={sFilePath};Extended Properties='Excel 12.0;'";
            using (var conn = new OleDbConnection(sConnectionString))
            {
                conn.Open();
                listSheetsName.Clear();
                var dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtSheetName == null) return;
                listSheetsName.AddRange(dtSheetName.Rows.Cast<DataRow>().Select(r => r[2].ToString()));//获取Excel的表名
                if (listSheetsName.Count <= 0) return;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select * from [" + listSheetsName[nIndex] + "]";
                    var ds = new DataSet();
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        da.Fill(ds, listSheetsName[nIndex]);
                        dtExcelData = ds.Tables[0];
                    }
                }
            }
        }

        #region 整理数据相关方法
        /// <summary>
        /// 整理联合要求
        /// </summary>
        /// <param name="dtExcelData"></param>
        /// <returns></returns>
        public static DataTable FormatDemand(DataTable dtExcelData)
        {
            var dtFormatData = dtExcelData.Copy();
            var sClientCodePattern = @"(?<=(\(|（))[0-9]{4}(?=(\)|）))";
            foreach (DataRow drData in dtExcelData.Rows)
            {
                if (drData["客户编号"].ToString() == drData["申请人编号"].ToString())
                {
                    dtFormatData.Rows.Cast<DataRow>().First(r => r.ItemArray.SequenceEqual(drData.ItemArray)).Delete();
                    dtFormatData.AcceptChanges();
                    var drCopyRow = dtFormatData.Rows.Add(drData.ItemArray);
                    drCopyRow["申请人编号"] = "";
                    var drCopyRow1 = dtFormatData.Rows.Add(drData.ItemArray);
                    drCopyRow1["客户编号"] = "";
                }
                else
                {
                    var mathces = Regex.Matches(drData["描述"].ToString(), sClientCodePattern);
                    var bIsUnionDemand = false;
                    foreach (var sValue in mathces.Cast<Match>().Select(m => m.Value).Distinct())
                    {
                        if (!string.IsNullOrWhiteSpace(drData["客户编号"].ToString()))
                        {
                            if (sValue != drData["客户编号"].ToString())
                            {
                                var drCopyRow = dtFormatData.Rows.Add(drData.ItemArray);
                                drCopyRow["是否联合要求"] = "Y";
                                drCopyRow["申请人编号"] = sValue;
                                bIsUnionDemand = true;
                            }
                        }
                        else if (!string.IsNullOrWhiteSpace(drData["申请人编号"].ToString()))
                        {
                            if (sValue != drData["申请人编号"].ToString())
                            {
                                var drCopyRow = dtFormatData.Rows.Add(drData.ItemArray);
                                drCopyRow["是否联合要求"] = "Y";
                                drCopyRow["客户编号"] = sValue;
                                bIsUnionDemand = true;
                            }
                        }
                    }
                    if (bIsUnionDemand)
                    {
                        dtFormatData.Rows.Cast<DataRow>().First(r => r.ItemArray.SequenceEqual(drData.ItemArray)).Delete();
                        dtFormatData.AcceptChanges();
                    }
                }
            }
            return dtFormatData;
        }

        /// <summary>
        /// 处理委托人实体
        /// </summary>
        /// <param name="dtExcelData"></param>
        /// <returns></returns>
        public static DataTable FormatEntity(DataTable dtExcelData)
        {
            var dtFormatData = dtExcelData.Copy();
            foreach (var dataRow in dtExcelData.Select("实体ID LIKE '%;%'"))
            {
                for (int i = 0; i < dataRow["实体ID"].ToString().Split(';').Length; i++)
                {
                    var sEntityID = dataRow["实体ID"].ToString().Split(';')[i];
                    if (string.IsNullOrEmpty(sEntityID)) continue;
                    var drCopyRow = dtFormatData.Rows.Add(dataRow.ItemArray);
                    drCopyRow["实体ID"] = sEntityID;
                    drCopyRow["原名"] = dataRow["原名"].ToString().Split(';').Length > i
                        ? dataRow["原名"].ToString().Split(';')[i]
                        : string.Empty;
                    drCopyRow["译名"] = dataRow["译名"].ToString().Split(';').Length > i
                        ? dataRow["译名"].ToString().Split(';')[i]
                        : string.Empty;
                }
            }
            dtFormatData.Select("实体ID LIKE '%;%'").ToList().ForEach(r => r.Delete());
            return dtFormatData;
        }

        /// <summary>
        /// 提取卷号中的类型代码
        /// </summary>
        /// <param name="dtExcelData"></param>
        /// <returns></returns>
        public static DataTable GetCaseType(DataTable dtExcelData)
        {
            var dtFormatData = dtExcelData.Copy();
            dtFormatData.Columns.Add("类型", typeof(string));
            var sCaseNoPattern = @"(?<=\d{2})[A-z]+?(?=\d+)";
            foreach (DataRow dataRow in dtFormatData.Rows)
            {
                var match = Regex.Match(dataRow["卷号"].ToString(), sCaseNoPattern);
                if (match.Success)
                    dataRow["类型"] = match.Value;
            }
            return dtFormatData;
        }

        /// <summary>
        /// 将多条同一案件的发明人信息合并为一条
        /// </summary>
        /// <param name="dtExcelData"></param>
        /// <returns></returns>
        public static DataTable MergeInventorInfo(DataTable dtExcelData)
        {
            var dtFormatData = dtExcelData.Copy();
            foreach (var groupInfo in dtExcelData.Rows.Cast<DataRow>().GroupBy(d => d["我方卷号"]).Where(g => g.Count() > 1))
            {
                foreach (var dataRow in dtFormatData.Select($"我方卷号 = '{groupInfo.Key}'"))
                {
                    dataRow.Delete();
                }
                var newRow = dtFormatData.NewRow();
                for (var i = 0; i < newRow.ItemArray.Length; i++)
                {
                    var listValue =
                        groupInfo.ToList().OrderBy(g => g["PID"]).Select(d => d.ItemArray[i].ToString()).ToList();
                    newRow[i] = string.Join(";", listValue);
                }
                newRow["我方卷号"] = groupInfo.Key;
                dtFormatData.Rows.Add(newRow);
            }
            return dtFormatData;
        }
        #endregion
    }
}
