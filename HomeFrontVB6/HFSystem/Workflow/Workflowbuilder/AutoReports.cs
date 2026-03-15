using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using C1.Win.C1FlexGrid;


namespace AutoNotice
{
    public partial class AutoReports : Form
    {
        public AutoReports()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Connection String
        /// </summary>
        private string sqlConnStr;

        /// <summary>
        /// 
        /// </summary>
        private bool IsModified;

        // Connection string
        public string ConnectionString
        {
            get
            {
                return sqlConnStr;
            }
            set
            {
                sqlConnStr = value;
            }
        }

        private void AutoReports_Load(object sender, EventArgs e)
        {
            this.sqlConn.ConnectionString = sqlConnStr;
            this.sqlDAAutoReports.Fill(dsAutoReports);

            SetupAutoReportsList();
            PopulateAutoReportsList();

            IsModified = false;
            flex.Focus();
        }

         /// <summary>
        /// Setup AutoReports list
        /// </summary>
        private void SetupAutoReportsList()
        {
            flex.Rows.Count = 1;
            flex.Cols.Count = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Columns.Count + 1;
            flex.ExtendLastCol = true;

            // create styles with data types, formats, etc
            CellStyle cs = flex.Styles.Add("DayOfWeek");
            cs.DataType = typeof(string);
            cs.ComboList = "|Every Day|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday";

            CellStyle cs1 = flex.Styles.Add("DayOfMonth");
            cs1.DataType = typeof(string);
            cs1.ComboList = "|Last Day|1|2|3|4|5|6|7|*|(|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31";

            CellStyle cs2 = flex.Styles.Add("EveryNthWeek");
            cs2.DataType = typeof(string);
            cs2.ComboList = "|1|2|3|4|5";

            CellStyle cs3 = flex.Styles.Add("EveryNthMonth");
            cs3.DataType = typeof(string);
            cs3.ComboList = "|1|2|3|4|5|6|7|*|(|10|11|12";

        }

        /// <summary>
        /// Converts Frequency value string value to number.
        /// </summary>
        /// <param name="freqType"></param>
        /// <param name="freqValue"></param>
        /// <returns></returns>
        private int ConvertFrequencyValueToNumber(string freqType,string freqValue)
        {
            int result = -1;
            switch (freqType)
            {
                case "Day of week":
                    {
                        switch (freqValue)
                        {
                            case "Every Day":
                                {
                                    result = 0;
                                    break;
                                }
                            case "Sunday":
                                {
                                    result = 1;
                                    break;
                                }
                            case "Monday":
                                {
                                    result = 2;
                                    break;
                                }
                            case "Tuesday":
                                {
                                    result = 3;
                                    break;
                                }
                            case "Wednesday":
                                {
                                    result = 4;
                                    break;
                                }
                            case "Thursday":
                                {
                                    result = 5;
                                    break;
                                }
                            case "Friday":
                                {
                                    result = 6;
                                    break;
                                }
                            case "Saturday":
                                {
                                    result = 7;
                                    break;
                                }
                        }
                        break;
                    }
                case "Day of month":
                    {
                        result = (freqValue == "Last Day") ? 0 : Convert.ToInt32(freqValue);
                        break;
                    }
                case "Every nth month":
                    {
                        result = Convert.ToInt32(freqValue);
                        break;
                    }
                case "Every nth week":
                    {
                        result = (freqValue == "Last Week") ? 0 : Convert.ToInt32(freqValue);
                        break;
                    }
            }
            return result;
        }

        /// <summary>
        /// Converts Frequency value numeric value to string.
        /// </summary>
        /// <param name="freqType"></param>
        /// <param name="freqValue"></param>
        /// <returns></returns>
        private string ConvertFrequencyValueToString(string freqType, int freqValue)
        {
            string result = "";
            switch (freqType)
            {
                case "Day of week":
                    {
                        switch (freqValue)
                        {
                            case 0:
                                {
                                    result = "Every Day";
                                    break;
                                }
                            case 1:
                                {
                                    result = "Sunday";
                                    break;
                                }
                            case 2:
                                {
                                    result = "Monday";
                                    break;
                                }
                            case 3:
                                {
                                    result = "Tuesday";
                                    break;
                                }
                            case 4:
                                {
                                    result = "Wednesday";
                                    break;
                                }
                            case 5:
                                {
                                    result = "Thursday";
                                    break;
                                }
                            case 6:
                                {
                                    result = "Friday";
                                    break;
                                }
                            case 7:
                                {
                                    result = "Saturday";
                                    break;
                                }
                        }
                        break;
                    }
                case "Day of month":
                    {
                        result = (freqValue == 0) ? "Last Day" : freqValue.ToString();
                        break;
                    }
                case "Every nth month":
                    {
                        result = freqValue.ToString();
                        break;
                    }
                case "Every nth week":
                    {
                        result = freqValue.ToString();
                        break;
                    }
            }
            return result;
        }

        private void SetFrequencyValueStyle(int rowNumb,string freqType)
        {
            switch (freqType)
            {
                case "Day of week":
                    {
                        CellRange rg = flex.GetCellRange(rowNumb, flex.Cols["FrequencyValue"].Index);
                        rg.Style = flex.Styles["DayOfWeek"];
                        break;
                    }
                case "Day of month":
                    {
                        CellRange rg = flex.GetCellRange(rowNumb, flex.Cols["FrequencyValue"].Index);
                        rg.Style = flex.Styles["DayOfMonth"];
                        break;
                    }
                case "Every nth month":
                    {
                        CellRange rg = flex.GetCellRange(rowNumb, flex.Cols["FrequencyValue"].Index);
                        rg.Style = flex.Styles["EveryNthMonth"];
                        break;
                    }
                case "Every nth week":
                    {
                        CellRange rg = flex.GetCellRange(rowNumb, flex.Cols["FrequencyValue"].Index);
                        rg.Style = flex.Styles["EveryNthWeek"];
                        break;
                    }
            }
        }

        /// <summary>
        /// Populates reports list
        /// </summary>
        private void PopulateAutoReportsList()
        {
            string name;
            flex.Rows.Count = 1;
            int row = 1;

            while (row <= dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows.Count)
            {
                flex.Rows.Add();
                for (int i = 1; i < flex.Cols.Count; i++)
                {
                    name = flex.Cols[i].Name;
                    switch(name)
                    {
                        case "TimeofDay":
                            {
                                flex[row, flex.Cols[name].Index] = Convert.ToDateTime(dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[row - 1][name]).ToShortTimeString();
                                break;
                            }
                        case "NextRun":
                            {
                                flex[row, flex.Cols[name].Index] = Convert.ToDateTime(dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[row - 1][name]).ToShortDateString();
                                break;
                            }
                        case "FrequencyValue":
                            {
                                flex[row, flex.Cols[name].Index] = ConvertFrequencyValueToString(dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[row - 1]["FrequencyType"].ToString(),
                                                Convert.ToInt32(dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[row - 1][name]));
                                SetFrequencyValueStyle(row, flex[row,flex.Cols["FrequencyType"].Index].ToString());
                                break;
                            }
                        default:
                            {
                                flex[row, flex.Cols[name].Index] = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[row - 1][name];
                                break;
                            }
                    }
                }
                row = flex.Rows.Count;
            }
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdDelete_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Are you sure that you want to delete selected reports ?",
                                 "Deleting reports", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.No) return;

            for (int i = flex.Selection.r1; i <= flex.Selection.r2; i++)
            {
                string ID = flex[i, flex.Cols["Seq"].Index].ToString();
                DeleteFromDataset(ID);
            }
            flex.Rows.RemoveRange(flex.Selection.r1, flex.Selection.r2 - flex.Selection.r1 + 1);
            IsModified = true;
        }

        /// <summary>
        /// Deletes record from dataset
        /// </summary>
        /// <param name="UserID"></param>
        private void DeleteFromDataset(string SeqID)
        {
            for (int i = 0; i < dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows.Count; i++)
            {
                DataRow row = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[i];
                if (row.RowState != DataRowState.Deleted)
                {
                    if ((row["Seq"].ToString() == SeqID)
                        && (row.RowState != DataRowState.Deleted))
                    {
                        row.Delete();
                        return;
                    }
                }
            }

        }

        private void AutoReports_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsModified)
            {
                DialogResult res = MessageBox.Show("Data have been modified. Would you like to save ?",
                                 "Saving", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    //if (!IsValidated())
                    //{
                    //    e.Cancel = true;
                    //    return;
                    //}
                    Save();
                }
            }
        }

        private void InsertRowIntoDataset(int rowNumb)
        {           
            DataRow row = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].NewRow();
            row["Description"] = flex[rowNumb, "Description"];
            row["RecipientAddress"] = flex[rowNumb, "RecipientAddress"] ;
            row["FromAddress"] = flex[rowNumb, "FromAddress"] ;
            row["Subject"] = flex[rowNumb, "Subject"];
            row["Message"] = flex[rowNumb, "Message"];
            row["Attachment"] = flex[rowNumb, "Attachment"];
            row["Report"] = flex[rowNumb, "Report"];
            row["TimeofDay"] = flex[rowNumb, "TimeofDay"];
            row["NextRun"] = flex[rowNumb, "NextRun"];
            row["FrequencyType"] = flex[rowNumb, "FrequencyType"];
            row["FrequencyValue"] = ConvertFrequencyValueToNumber(flex[rowNumb, "FrequencyType"].ToString(), 
                                                flex[rowNumb, "FrequencyValue"].ToString());
            row["Param1"] = flex[rowNumb, "Param1"];
            row["Param2"] = flex[rowNumb, "Param2"];
            row["Param3"] = flex[rowNumb, "Param3"];
            row["Param4"] = flex[rowNumb, "Param4"];
            row["Param5"] = flex[rowNumb, "Param5"];
            row["Param6"] = flex[rowNumb, "Param6"];
            row["Param7"] = flex[rowNumb, "Param7"];
            row["Param8"] = flex[rowNumb, "Param8"];
            row["Param9"] = flex[rowNumb, "Param9"];
            row["Param10"] = flex[rowNumb, "Param10"];
            dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows.Add(row);

            flex[rowNumb, "Seq"] = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows[dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows.Count-1]["Seq"];
        }

        /// <summary>
        /// Saves data from the grid to the dataset
        /// </summary>
        private void SaveToDataset()
        {
            string name;
            for (int i = 1; i < flex.Rows.Count; i++)
            {
                DataRow row = dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Rows.Find(flex[i, flex.Cols["Seq"].Index]);

                if (row == null)
                {
                    InsertRowIntoDataset(i);
                    return;
                }
                
                for (int j = 2; j < flex.Cols.Count; j++)
                {
                    name = flex.Cols[j].Name;
                    if (name == "FrequencyValue")
                        row[name] = ConvertFrequencyValueToNumber(flex[i, flex.Cols["FrequencyType"].Index].ToString(),flex[i, flex.Cols[name].Index].ToString());
                    else if ((flex[i, flex.Cols[name].Index] != null)
                        && ((flex[i, flex.Cols[name].Index].ToString() != "")))
                        row[name] = flex[i, flex.Cols[name].Index];
                    else if (dsAutoReports.Tables["AutoNotice_GetAutoReportsList"].Columns[name].AllowDBNull)
                        row[name] = Convert.DBNull;
                    else
                        row[name] = false;
                }
            }
        }

        private void SaveToDatabase()
        {
            try
            {
                this.sqlDAAutoReports.Update(dsAutoReports);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Saving reports", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        /// <summary>
        /// Saves data
        /// </summary>
        private void Save()
        {
            //if (!IsValidated()) return;
            SaveToDataset();
            SaveToDatabase();
            IsModified = false;
        }

        private void cmdNew_Click(object sender, EventArgs e)
        {
            AutoReportDetails details = new AutoReportDetails();
            details.ReportID = -1;
            details.ConnectionString = sqlConnStr;
            details.ShowDialog(this);

            if (details.Canceled) return;

            flex.Rows.Add();
            flex.Row = flex.Rows.Count - 1;
            int row = flex.Row;

            flex[row, "Seq"] = 0;
            flex[row, "Description"] = details.Description;
            flex[row, "RecipientAddress"] = details.Recipient;
            flex[row, "FromAddress"] = details.Sender;
            flex[row, "Subject"] = details.Subject;
            flex[row, "Message"] = details.Message;
            flex[row, "Attachment"] = details.Attachment;
            flex[row, "Report"] = details.Report;

            flex[row, "TimeofDay"] = details.RunTime.ToShortTimeString();
            flex[row, "NextRun"] = details.NextRun.ToShortDateString(); ;

            flex[row, "FrequencyType"] = details.FrequencyType;
            flex[row, "FrequencyValue"] = ConvertFrequencyValueToString(details.FrequencyType, details.FrequencyValue);

            if (!string.IsNullOrEmpty(details.Report))
            {
                // get report parameters
                ArrayList rep_params = details.ReportParameters;
                for (int i = 0; i < rep_params.Count; i++)
                {
                    ReportParam param = (ReportParam)rep_params[i];
                    flex[row, "Param" + (param.Index + 1)] = param.Value;
                }
            }
            IsModified = true;
        }

        private void flex_DoubleClick(object sender, EventArgs e)
        {
            OnEdit();
        }

        private void OnEdit()
        {
            HitTestInfo info = flex.HitTest();
            if (info.Type == HitTestTypeEnum.ColumnHeader) return;

            int row = flex.Row;

            AutoReportDetails details = new AutoReportDetails();
            details.ReportID = Convert.ToInt32(flex[row, "Seq"]);
            details.Description = flex[row, "Description"].ToString();
            details.Recipient = flex[row, "RecipientAddress"].ToString();
            details.Sender = flex[row, "FromAddress"].ToString();
            details.Subject = flex[row, "Subject"].ToString();
            details.Message = flex[row, "Message"].ToString();
            details.Attachment = flex[row, "Attachment"].ToString();
            details.Report = flex[row, "Report"].ToString();

            details.RunTime = Convert.ToDateTime(flex[row, "TimeofDay"]);
            details.NextRun = Convert.ToDateTime(flex[row, "NextRun"]);

            details.FrequencyType = flex[row, "FrequencyType"].ToString();
            details.FrequencyValue = ConvertFrequencyValueToNumber(flex[row, "FrequencyType"].ToString(),
                                    flex[row, "FrequencyValue"].ToString());
            details.RecipientQuery = flex[row, "RecipientQuery"].ToString();

            //if (!string.IsNullOrEmpty(details.Report))
            //{
                // set report parameters
                ArrayList rep_params = new ArrayList();
                for (int i = 1; i <= 10; i++)
                {
                    ReportParam param = new ReportParam();
                    param.Name = "Param" + i;
                    param.Index = i - 1;
                    param.Value = (flex[row, "Param" + i] != null)
                                ? flex[row, "Param" + i].ToString()
                                : "";
                    rep_params.Add(param);
                }
                details.ReportParameters = rep_params;
           // }

            details.ConnectionString = sqlConnStr;
            details.ShowDialog(this);

            if ((details.Canceled)
                || (!details.Modified))
            {
                flex.AllowEditing = true;
                return;
            }

            flex[row, "Description"] = details.Description;
            flex[row, "RecipientAddress"] = details.Recipient;
            flex[row, "FromAddress"] = details.Sender;
            flex[row, "Subject"] = details.Subject;
            flex[row, "Message"] = details.Message;
            flex[row, "Attachment"] = details.Attachment;
            flex[row, "Report"] = details.Report;

            flex[row, "TimeofDay"] = details.RunTime.ToShortTimeString();
            flex[row, "NextRun"] = details.NextRun.ToShortDateString();

            flex[row, "FrequencyType"] = details.FrequencyType;
            SetFrequencyValueStyle(row, flex[row, "FrequencyType"].ToString());   
            flex[row, "FrequencyValue"] = ConvertFrequencyValueToString(details.FrequencyType, details.FrequencyValue);
            flex[row, "RecipientQuery"] = details.RecipientQuery;

            //if (!string.IsNullOrEmpty(details.Report))
            //{
                // get report parameters
                ArrayList report_params = details.ReportParameters;
                for (int i = 0; i < report_params.Count; i++)
                {
                    ReportParam param = (ReportParam)report_params[i];
                    flex[row, "Param" + (param.Index + 1)] = param.Value;
                }
            //}

            // format subject and message
            if (!string.IsNullOrEmpty(details.Subject))
            {
                string formattedSubject = details.Subject.Replace("'", "''''");
                formattedSubject = formattedSubject.Replace("D{", "ISNULL(CONVERT(VARCHAR(12),");
                formattedSubject = formattedSubject.Replace("}D", ",109),'')");

                formattedSubject = formattedSubject.Replace("N{", "ISNULL(CAST(CAST(round(");
                formattedSubject = formattedSubject.Replace("}N", ",2) AS money) AS varchar),'')");

                formattedSubject = formattedSubject.Replace("{{", "'+");
                formattedSubject = formattedSubject.Replace("}}", "+'");
                formattedSubject = "'" + formattedSubject + "'";
                flex[row, "FormattedSubject"] = formattedSubject;
            }

            if (!string.IsNullOrEmpty(details.Message))
            {
                string formattedMessage = details.Message.Replace("'", "''''");
                formattedMessage = formattedMessage.Replace("D{", "ISNULL(CONVERT(VARCHAR(12),");
                formattedMessage = formattedMessage.Replace("}D", ",109),'')");

                formattedMessage = formattedMessage.Replace("N{", "ISNULL(CAST(CAST(round(");
                formattedMessage = formattedMessage.Replace("}N", ",2) AS money) AS varchar),'')");

                formattedMessage = formattedMessage.Replace("{{", "'+");
                formattedMessage = formattedMessage.Replace("}}", "+'");
                formattedMessage = "'" + formattedMessage + "'";
                flex[row, "FormattedMessage"] = formattedMessage;
            }

            flex.AllowEditing = true;
            IsModified = true;
        }

        private void flex_BeforeDoubleClick(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            flex.AllowEditing = false;
        }

        private void flex_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (flex.Cols[e.Col].Name == "FrequencyType")
            {
                SetFrequencyValueStyle(e.Row, flex[e.Row, e.Col].ToString());         
            }
            IsModified = true;
        }

        private void flex_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                OnEdit();
        }

    }
}
