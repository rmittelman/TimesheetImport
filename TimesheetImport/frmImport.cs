using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Aimm.Logging;
using System.Windows.Forms;
using System.Text;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;

namespace TimesheetImport
{
    public partial class frmImport : Form
    {
        public frmImport()
        {
            InitializeComponent();
        }

        #region enums

        enum cols
        {
            empName = 1,
            empID = 2,
            workOrder = 3,
            custName = 4,
            jobNo = 5,
            startTime = 6,
            stopTime = 7,
            hours = 8,
            regHours = 9,
            otHours = 10,
            dblOtHours = 11,
            status = 12,
            message = 13
        }

        enum cellColors
        {
            errorColor = 3,
            warnColor = 55
        }

        #endregion

        #region objects

        clsExcel oXl = null;
        dynamic xlRange = null;
        dynamic xlCell = null;
        ToolTip toolTip1 = new ToolTip();

        #endregion

        #region variables

        string connString = Properties.Settings.Default.POLSQL;
        string sourcePath = Properties.Settings.Default.SourceFolder;
        string archivePath = Properties.Settings.Default.ArchiveFolder;
        string errorPath = Properties.Settings.Default.ErrorFolder;
        string logPath = Properties.Settings.Default.LogFolder;
        bool showExcel = (bool)Properties.Settings.Default.ShowExcel;
        string excelRange = Properties.Settings.Default.ExcelRange;

        bool isValid = false;
        bool allValid = true;
        string xlPathName = "";
        string xlFile = "";
        string destPath = "";
        string destFile = "";
        string destPathName = "";
        string logFile = "TimesheetImport.log";
        string logPathName = "";

        #endregion

        #region properties

        private string _status;
        public string Status
        {
            set
            {
                _status = value;
                txtStatus.Text = value;
            }
            get { return _status; }
        }

        #endregion

        #region events

        private void frmImport_Load(object sender, EventArgs e)
        {
            // set tooltips
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.SetToolTip(btnFindExcel, "Find Excel Timesheets File");
            toolTip1.SetToolTip(btnImport, "Import Timesheets from Excel File");


        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            btnImport.Enabled = false;
            string msg = "";
            List<string> timesheetsEntered = new List<string>();

            // continue if we can open excel file
            xlPathName = txtExcelFile.Text;
            if(open_excel(xlPathName))
            {
                xlFile = Path.GetFileName(xlPathName);
                msg = $"Opened Excel file \"{xlFile}\"";
                Status = msg;
                LogIt.LogInfo(msg);

                // if csv file, save-as excel file.
                if(save_csv_as_excel())
                {
                    // get and resize range
                    isValid = oXl.GetRange(excelRange);
                    if(isValid)
                    {
                        var activeRows = oXl.ActiveRows;
                        isValid = oXl.ResizeRange(rows: activeRows);
                    }

                    if(isValid)
                    {
                        msg = "Identified active timesheets range";
                        Status = msg;
                        LogIt.LogInfo(msg);

                        // add headings for last 2 columns
                        oXl.WorkSheet.Range("$L$1").Value = "Status";
                        oXl.WorkSheet.Range("$M$1").Value = "Message";

                        // resize range
                        oXl.Range.Columns.AutoFit();

                        // start processing the file
                        allValid = true;

                        // loop thru each invoice row on worksheet
                        foreach(dynamic xlRow in oXl.Range.Rows)
                        {
                            isValid = false;
                            string timesheetId = "";
                            DateTime startTime = new DateTime();
                            DateTime stopTime = new DateTime();

                            try
                            {
                                string empName = (xlRow.Cells[cols.empName].Value ?? "").ToString().Trim();
                                string workOrder = (xlRow.Cells[cols.workOrder].Value ?? "").ToString().Trim();
                                string jobNo = (xlRow.Cells[cols.jobNo].Value ?? "").ToString().Trim();
                                int jobId = 0;
                                int empId = 0;
                                float regHours = 0;
                                float otHours = 0;
                                float dblOtHours = 0;

                                // validate employee id is number
                                xlCell = xlRow.Cells[cols.empID];
                                if(int.TryParse((xlCell.Value ?? "").ToString().Trim(), out empId))
                                {
                                    // validate start, stop times
                                    bool validStart = DateTime.TryParse((xlRow.Cells[cols.startTime].Value ?? "").ToString().Trim(), out startTime);
                                    bool validStop = DateTime.TryParse((xlRow.Cells[cols.stopTime].Value ?? "").ToString().Trim(), out stopTime);
                                    bool sameDay = true;
                                    if(validStart && validStop)
                                        sameDay = (startTime.Date == stopTime.Date);

                                    if(validStart && validStop && sameDay)
                                    {
                                        // validate job id is numeric and in AIMM
                                        bool validJob = int.TryParse(jobNo, out jobId);
                                        if(validJob)
                                            validJob = valid_job(jobId, connString);

                                        if(validJob)
                                        {
                                            // validate work order belongs to job
                                            bool validWorkOrder = valid_work_order(jobId, workOrder, connString);
                                            if(validWorkOrder)
                                            {
                                                // validate hours
                                                // ignore total hours, it includes 1/2 hour for lunch.
                                                bool validReg = float.TryParse((xlRow.Cells[cols.regHours].Value ?? "").ToString().Trim(), out regHours);
                                                bool validOt = float.TryParse((xlRow.Cells[cols.otHours].Value ?? "").ToString().Trim(), out otHours);
                                                bool validDot = float.TryParse((xlRow.Cells[cols.dblOtHours].Value ?? "").ToString().Trim(), out dblOtHours);
                                                if(validReg && validOt && validDot)
                                                {


                                                    // get or create weekly timesheet record for employee
                                                    timesheetId = get_timesheet(startTime, empId, connString);
                                                    if(timesheetId == "")
                                                        timesheetId = create_timesheet(startTime, empId, connString);

                                                    // add detail for employee
                                                    isValid = add_timesheet_detail(timesheetId, jobId, workOrder, startTime, regHours, otHours, dblOtHours, connString);

                                                    // save timesheet IDs so we can update them at the end
                                                    if(isValid)
                                                    {
                                                        if(!timesheetsEntered.Contains(timesheetId))
                                                        {
                                                            timesheetsEntered.Add(timesheetId);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        msg = $"Could not add timesheet for {empName} for {startTime.ToShortDateString()} for job {jobNo}";
                                                        Status = msg;
                                                        LogIt.LogError(msg);
                                                    }
                                                }
                                                else
                                                {
                                                    msg = $"Timesheet for {empName} for job {jobNo} has bad hours";
                                                    Status = msg;
                                                    LogIt.LogError(msg);
                                                    if(!validReg)
                                                        color_excel_cell(xlRow.Cells[cols.regHours], cellColors.errorColor);
                                                    if(!validOt)
                                                        color_excel_cell(xlRow.Cells[cols.otHours], cellColors.errorColor);
                                                    if(!validDot)
                                                        color_excel_cell(xlRow.Cells[cols.dblOtHours], cellColors.errorColor);
                                                    set_excel_status(xlRow, "Error", "Bad hours");
                                                }

                                            }
                                            else
                                            {
                                                msg = $"Timesheet for {empName} for job {jobNo} has invalid work order number: {workOrder}";
                                                Status = msg;
                                                LogIt.LogError(msg);
                                                color_excel_cell(xlRow.Cells[cols.workOrder], cellColors.errorColor);
                                            }

                                        }
                                        else
                                        {
                                            msg = $"Timesheet for {empName} has invalid job ID: {jobNo}";
                                            Status = msg;
                                            LogIt.LogError(msg);
                                            color_excel_cell(xlRow.Cells[cols.jobNo], cellColors.errorColor);
                                        }

                                    }
                                    else
                                    {
                                        msg = $"Timesheet for {empName} for job {jobNo} has bad time(s) or dates don't agree";
                                        Status = msg;
                                        LogIt.LogError(msg);
                                        if(!validStart)
                                            color_excel_cell(xlRow.Cells[cols.startTime], cellColors.errorColor);

                                        if(!validStop)
                                            color_excel_cell(xlRow.Cells[cols.stopTime], cellColors.errorColor);

                                        if(!sameDay)
                                        {
                                            color_excel_cell(xlRow.Cells[cols.startTime], cellColors.errorColor);
                                            color_excel_cell(xlRow.Cells[cols.stopTime], cellColors.errorColor);
                                        }

                                        set_excel_status(xlRow, "Error", "Bad date(s) or dates don't agree");
                                    } // valid start, stop times

                                }
                                else
                                {
                                    isValid = false;
                                    msg = $"Timesheet for {empName} for job {jobNo} has bad employee ID";
                                    Status = msg;
                                    LogIt.LogError(msg);
                                    color_excel_cell(xlCell, cellColors.errorColor);
                                    set_excel_status(xlRow, "Error", "Bad employee ID");
                                } // valid employee ID

                            }
                            catch(Exception ex)
                            {
                                msg = $"Error processing timesheet {timesheetId} for {startTime.ToShortDateString()}: {ex.Message}";
                                LogIt.LogError(msg);
                                Status = msg;
                                color_excel_cell(xlRow.Cells[cols.empName], cellColors.errorColor);
                                set_excel_status(xlRow, "Error", ex.Message);
                            }

                            // keep track if all items were valid
                            allValid = allValid && isValid;
                        }

                        // update weekly timesheet totals for each timesheet id in list
                        isValid = update_timesheet_totals(timesheetsEntered, connString);
                        if(isValid)
                        {
                            msg = "Timesheets imported and totals updated";
                            Status = msg;
                            LogIt.LogInfo(msg);
                        }

                        // save excel file if any invalid items.
                        isValid = oXl.CloseWorkbook(!allValid);

                        // move workbook to archive/errors folder
                        destPath = allValid ? archivePath : errorPath;
                        destFile = string.Concat(
                            Path.GetFileNameWithoutExtension(xlFile),
                            DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"),
                            Path.GetExtension(xlFile));
                        destPathName = Path.Combine(destPath, destFile);
                        if(move_file(xlPathName, destPathName))
                        {
                            txtExcelFile.Text = destPathName;
                            if(allValid)
                            {
                                msg = $"Import completed without errors. Moved \"{xlFile}\" to \"{destPathName}\"";
                                LogIt.LogInfo(msg);
                            }
                            else
                            {
                                msg = $"Import completed with errors. Moved \"{xlFile}\" to \"{destPathName}\"";
                                LogIt.LogWarn(msg);
                            }
                            Status = msg;
                        }

                        oXl.CloseExcel();
                        oXl = null;

                    }
                    else
                    {
                        msg = "Could not identify active timesheets range, timesheets not imported.";
                        Status = msg;
                        LogIt.LogError(msg);
                    } // got active timesheets

                }
                else
                {
                    msg = $"Could not save {xlFile} as standard Excel workbook, timesheets not imported.";
                    Status = msg;
                    LogIt.LogError(msg);
                }

            }
            else
            {
                msg = $"Could not open Excel file \"{xlPathName}\", timesheets not imported";
                LogIt.LogError(msg);
                Status = msg;
            }
        }

        /// <summary>
        /// Find file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFindExcel_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if(btn != null)
            {
                using(OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.InitialDirectory = sourcePath;
                    ofd.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx, *.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                    ofd.FilterIndex = 1;
                    if(ofd.ShowDialog() == DialogResult.OK)
                    {
                        txtExcelFile.Text = ofd.FileName;
                    }
                }
            }

        }

        #endregion

        #region methods

        private bool move_file(string sourcePath, string destPath)
        {
            try
            {
                File.Move(sourcePath, destPath);
                return true;
            }
            catch(Exception ex)
            {
                var msg = $"Error moving file \"{Path.GetFileName(sourcePath)}\" to \"{Path.GetDirectoryName(sourcePath)}\": {ex.Message}";
                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                LogIt.LogError(msg);
                return false;
            }
        }

        /// <summary>
        /// start ms excel and open supplied workbook name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>boolean indicating success status</returns>
        private bool open_excel(string fileName)
        {
            bool result = false;
            string msg = "";

            if(File.Exists(fileName))
            {
                var xlFile = Path.GetFileName(fileName);
                try
                {
                    oXl = new clsExcel();
                    oXl.Visible = showExcel;
                    result = oXl.OpenExcel(xlPathName);
                }
                catch(Exception ex)
                {
                    msg = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    LogIt.LogError(msg);
                }
            }
            return result;
        }

        private void set_excel_status(dynamic row, string status, string message)
        {
            if(row != null)
            {
                row.Cells[cols.status].Value = status;
                string txt = (row.Cells[cols.message].Value ?? "").ToString().Trim();
                if(txt != "")
                {
                    row.Cells[cols.message] = $"{txt}\n{message}";
                    row.Cells[cols.message].Style.WrapText = true;
                    row.EntireRow.AutoFit();
                }
                else
                    row.Cells[cols.message].Value = message;
                row.Cells[cols.message].Style.WrapText = true;
                row.EntireRow.AutoFit();
            }
        }

        private void color_excel_cell(dynamic cell, cellColors color)
        {
            cell.Interior.ColorIndex = color;
        }

        /// <summary>
        /// saves CSV file as XSLX file if needed
        /// </summary>
        /// <returns>boolean indicating file was already an excel file or was properly saved as excel file</returns>
        private bool save_csv_as_excel()
        {
            bool response = !xlPathName.EndsWith(".csv");
            if(!response)
            {
                var newPathName = xlPathName.Replace(".csv", ".xlsx");
                response = oXl.SaveWorkbookAs(newPathName);
                if(response)
                {
                    xlPathName = newPathName;
                    xlFile = Path.GetFileName(xlPathName);
                    LogIt.LogInfo($"Saved CSV file as {xlFile}");
                }

            }
            return response;
        }

        /// <summary>
        /// close excel file, save if needed, kill objects
        /// </summary>
        /// <param name="needToSave"></param>
        /// <returns>boolean indicating success status</returns>
        private bool close_excel(bool needToSave = false)
        {
            try
            {
                // close workbook, cleanup excel
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // close and release

                try
                {
                    oXl.CloseWorkbook(needToSave);
                    //xlWorkbook.Close(needToSave);

                }
                catch(COMException ex)
                {
                    // ignore if already closed
                }

                // release com objects to fully kill excel process from running in the background
                //try
                //{
                //    Marshal.ReleaseComObject(xlCell);
                //}
                //catch(NullReferenceException ex)
                //{
                //    // ignore if not yet instantiated
                //}
                //Marshal.ReleaseComObject(xlRange);
                //Marshal.ReleaseComObject(xlWorksheet);
                //Marshal.ReleaseComObject(xlWorkbook);

                //// quit and release
                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
                LogIt.LogInfo($"Closed Excel file, save = {needToSave}");
                return true;
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error closing excel file: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// get crew number and name for the supplied employee ID
        /// </summary>
        /// <param name="employeeID"></param>
        /// <param name="connectionString"></param>
        /// <returns>KeyValuePair containing crew ID and name</returns>
        private KeyValuePair<int, string> get_employee_crew(int employeeID, string connectionString)
        {
            KeyValuePair<int, string> crewInfo = new KeyValuePair<int, string>();
            string msg = $"Getting crew info for employee {employeeID}";

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "SELECT c.CrewID, c.CrewName FROM MLG.POL.tblCrewMembers cm "
                                   + $"INNER JOIN MLG.POL.tblCrews c ON cm.CrewID = c.CrewID WHERE cm.EmployeeID = {employeeID}";

                    using(SqlDataAdapter dap = new SqlDataAdapter(cmdText, conn))
                    {
                        using(DataTable dt = new DataTable())
                        {
                            dap.Fill(dt);
                            if(dt.Rows.Count > 0)
                                crewInfo = new KeyValuePair<int, string>((int)dt.Rows[0][0], (string)dt.Rows[0][1]);
                        }
                    }
                }

            }
            catch(Exception ex)
            {
                msg = $"Error getting crew for employee {employeeID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }





            //SqlDataReader reader = null;
            //try
            //{
            //    using(SqlConnection conn = new SqlConnection(connectionString))
            //    {

            //        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
            //        {
            //            cmd.Parameters.AddWithValue("@employeeID", employeeID);
            //            conn.Open();
            //            reader = cmd.ExecuteReader();
            //            if(reader.HasRows)
            //            {
            //                while(reader.Read())
            //                {
            //                    var record = (IDataRecord)reader;
            //                    crewInfo = new KeyValuePair<int, string>((int)record[0], (string)record[1]);
            //                }
            //            }
            //        }
            //    }
            //}
            //catch(Exception ex)
            //{
            //    msg = $"Error getting crew for employee {employeeID}: {ex.Message}";
            //    Status = msg;
            //    LogIt.LogError(msg);
            //}
            //finally
            //{
            //    try
            //    {
            //        reader.Close();
            //        reader = null;
            //    }
            //    catch(Exception)
            //    {

            //    }
            //}
            return crewInfo;

        }

        /// <summary>
        /// get week start and end dates for supplied date
        /// </summary>
        /// <param name="workDate">date to use to get start and end dates for week</param>
        /// <param name="wkEndDay"><see cref="DayOfWeek"/> enum member identifying the week-ending day</param>
        /// <returns></returns>
        private KeyValuePair<DateTime, DateTime> get_week_start_and_end(DateTime workDate, DayOfWeek wkEndDay)
        {
            int dow = (int)workDate.DayOfWeek;
            int wed = (int)wkEndDay;
            int daysToEow = (wed - dow) >= 0 ? wed - dow : wed - dow + 7;
            return new KeyValuePair<DateTime, DateTime>(workDate.AddDays(daysToEow - 6), workDate.AddDays(daysToEow));
        }

        /// <summary>
        /// find the weekly timesheet for the employee and date supplied
        /// </summary>
        /// <param name="workDate">date work reported for</param>
        /// <param name="employeeID">ID of employee reporting work</param>
        /// <returns>unique timesheet ID for employee for the week</returns>
        private string get_timesheet(DateTime workDate, int employeeID, string connectionString)
        {
            string result = "";

            // get start and end dates given the date from the excel file.
            KeyValuePair<DateTime, DateTime> wkStartAndEnd = get_week_start_and_end(workDate, DayOfWeek.Wednesday);

            LogIt.LogInfo($"Getting timesheet for {workDate.ToShortDateString()} for employee {employeeID}");

            // lookup the timesheet
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string wkStart = wkStartAndEnd.Key.ToString("yyyy-MM-dd 00:00:00.000");
                    string wkEnd = wkStartAndEnd.Value.ToString("yyyy-MM-dd 00:00:00.000");
                    string cmdText = "SELECT WeeklyTimeSheetID FROM MLG.POL.tblWeeklyTimeSheets "
                                   + $"WHERE EmployeeID = {employeeID} AND StartDate = '{wkStart}' AND EndDate = '{wkEnd}'";

                    using(SqlDataAdapter dap = new SqlDataAdapter(cmdText, conn))
                    {
                        using(DataTable dt = new DataTable())
                        {
                            dap.Fill(dt);
                            if(dt.Rows.Count > 0)
                                result = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting timesheet for {workDate.ToShortDateString()} for employee {employeeID}: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// create a weekly timesheet for employee
        /// </summary>
        /// <param name="workDate">Date work reported for</param>
        /// <param name="employeeID">ID of employee reporting work</param>
        /// <param name="connectionString"></param>
        /// <returns>(yymmXXXX)</returns>
        private string create_timesheet(DateTime workDate, int employeeID, string connectionString)
        {
            string result = "";
            // get start and end dates given the date from the excel file.
            KeyValuePair<DateTime, DateTime> wkStartAndEnd = get_week_start_and_end(workDate, DayOfWeek.Wednesday);
            var stDate = wkStartAndEnd.Key.ToString("yyyy-MM-dd 00:00:00.000");
            var endDate = wkStartAndEnd.Value.ToString("yyyy-MM-dd 00:00:00.000");

            // get list of already existing timesheet ids for the month
            List<string> timesheetsUsedForMonth = get_timesheets_for_month(workDate, connectionString);

            // build a timesheet id which hasn't been used for the month
            string timesheetID = get_unique_timesheet_id(workDate, timesheetsUsedForMonth);

            // get the crew
            KeyValuePair<int, string> crewInfo = get_employee_crew(employeeID, connectionString);

            // add timesheet record
            string sql = "INSERT INTO MLG.POL.tblWeeklyTimeSheets (WeeklyTimeSheetID, EmployeeID, StartDate, EndDate, TotalHours, RegHours, OTHours, CrewID, CrewName) "
                       + $"VALUES ('{timesheetID}', {employeeID}, '{stDate}', '{endDate}', 0, 0, 0, {crewInfo.Key}, '{crewInfo.Value}')";

            var isOk = false;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    using(SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        var recs = cmd.ExecuteNonQuery();
                        isOk = (recs == 1);
                    }
                }

            }
            catch(Exception ex)
            {
                string msg = $"Error adding weekly timesheet for {employeeID} for week {wkStartAndEnd.Key.ToShortDateString()} - {wkStartAndEnd.Value.ToShortDateString()}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }

            if(isOk)
                result = timesheetID;

            return result;
        }

        /// <summary>
        /// get all timesheet IDs for the month supplied
        /// </summary>
        /// <param name="timesheetDate"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private List<string> get_timesheets_for_month(DateTime timesheetDate, string connectionString)
        {
            string yymm = timesheetDate.ToString("yyMM");
            string msg = $"Getting timesheet IDs for \"{yymm}\"";
            List<string> timesheets = new List<string>();

            try
            {
                // get data using DataTable
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = $"SELECT WeeklyTimeSheetID FROM MLG.POL.tblWeeklyTimeSheets WHERE WeeklyTimeSheetID like '{yymm}%'";
                    using(SqlDataAdapter dap = new SqlDataAdapter(cmdText, conn))
                    {
                        using(DataTable dt = new DataTable())
                        {
                            dap.Fill(dt);
                            foreach(DataRow row in dt.Rows)
                            {
                                timesheets.Add(row[0].ToString());
                            }
                        }
                    }

                }

                //// get data using DataReader
                //SqlDataReader reader = null;
                //using(SqlConnection conn = new SqlConnection(connectionString))
                //{
                //    string cmdText = "SELECT WeeklyTimeSheetID FROM MLG.POL.tblWeeklyTimeSheets WHERE WeeklyTimeSheetID like @yymm";
                //    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                //    {
                //        cmd.Parameters.AddWithValue("@yymm", yymm + "%");
                //        conn.Open();
                //        reader = cmd.ExecuteReader();
                //        if(reader.HasRows)
                //        {
                //            while(reader.Read())
                //            {
                //                var record = (IDataRecord)reader;
                //                timesheets.Add((string)record[0]);
                //            }
                //        }
                //    }
                //}

            }
            catch(Exception ex)
            {
                msg = $"Error getting timesheet IDs for \"{yymm}\": {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            //finally
            //{
            //    try
            //    {
            //        reader.Close();
            //        reader = null;
            //    }
            //    catch(Exception)
            //    {

            //    }
            //}
            return timesheets;

        }

        /// <summary>
        /// build random timesheet id that hasn't been used in the current month
        /// </summary>
        /// <param name="timesheetDate"></param>
        /// <param name="existingTimesheets"></param>
        /// <returns></returns>
        private string get_unique_timesheet_id(DateTime timesheetDate, List<string> existingTimesheets)
        {
            string result = "";
            string yymm = timesheetDate.ToString("yyMM");
            const string pool = "ABCDFGHJKLMNPQRSTVWXYZ";
            do
            {
                result = generate_random_string(8, yymm, pool);
            } while(existingTimesheets.Contains(result));

            return result;
        }

        /// <summary>
        /// adds a timesheet entry for supplied timesheet ID
        /// </summary>
        /// <param name="timesheetId">8-digit hashed ID of timesheet to add hours to</param>
        /// <param name="jobId">AIMM job number</param>
        /// <param name="workOrder">AIMM work order number</param>
        /// <param name="startTime">Date-time work reported for</param>
        /// <param name="regularHours">Number of regular hours reported</param>
        /// <param name="overtimeHours">Number of overtime hours reported</param>
        /// <param name="doubleOvertimeHours">Number of double overtime hours reported</param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool add_timesheet_detail(string timesheetId, int jobId, string workOrder, DateTime startTime, float regularHours, float overtimeHours, float doubleOvertimeHours, string connectionString)
        {
            bool result = false;
            string estID = $"{jobId}-{workOrder}";
            float hours = regularHours + overtimeHours + doubleOvertimeHours;
            KeyValuePair<DateTime, DateTime> wkStartAndEnd = get_week_start_and_end(startTime, DayOfWeek.Wednesday);
            string workDate = startTime.Date.ToString("yyyy-MM-dd 00:00:00.000");
            string weDate = wkStartAndEnd.Value.ToString("yyyy-MM-dd 00:00:00.000");
            ;
            string sql = "INSERT INTO MLG.POL.tblWeeklyTimeSheetEntries (WeeklyTimeSheetID, JobID, HoursWorked, EstimateID, EstimateTypeID, TheDate, "
                       + "Notes, PerformanceLevel, AttendCode, NPReasonCode, Correction, JobErrorID, WeekEndingDate) VALUES "
                       + $"( '{timesheetId}', {jobId}, {hours}, '{estID}', null, '{workDate}', null, null, null, null, 0, null, '{weDate}')";

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    using(SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        var recs = cmd.ExecuteNonQuery();
                        result = (recs == 1);
                    }
                }
            }
            catch(Exception ex)
            {
                string msg = $"Error adding timesheet detail for timesheet \"{timesheetId}\" for {startTime.ToShortDateString()}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// updates timesheet totals for supplied timesheet IDs
        /// </summary>
        /// <param name="timesheetsUpdated">List of timesheet IDs to be updated</param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool update_timesheet_totals(List<string> timesheetsUpdated, string connectionString)
        {
            bool result = false;
            string where = string.Join("','",timesheetsUpdated);
            string sql = "UPDATE ts SET ts.TotalHours = tse.HoursWorked, ts.RegHours = case when tse.HoursWorked <= 40 then tse.HoursWorked else 40 end, "
                       + "ts.OTHours = case when tse.HoursWorked <= 40 then 0 else tse.HoursWorked - 40 end "
                       + "FROM MLG.POL.tblWeeklyTimeSheets ts INNER JOIN "
                       + "(SELECT WeeklyTimeSheetID, sum(HoursWorked) as HoursWorked FROM MLG.POL.tblWeeklyTimeSheetEntries GROUP BY WeeklyTimeSheetID) as tse "
                       + $"ON tse.WeeklyTimeSheetID = ts.WeeklyTimeSheetID WHERE ts.WeeklyTimeSheetID in('{where}')";

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    using(SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        var recs = cmd.ExecuteNonQuery();
                        result = (recs >= 1);
                    }
                }
            }
            catch(Exception ex)
            {
                string msg = $"Error updating timesheet totals: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// verify valid job
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="connectionString"></param>
        /// <returns>boolean indicating job exists in database</returns>
        private static bool valid_job(int jobID, string connectionString)
        {
            bool isValid = false;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Validating job ID {jobID}");
                    string cmdText = "SELECT COUNT(JobID) FROM MLG.POL.tblJobs WHERE JobID = @jobID";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        conn.Open();
                        int rows = (int)cmd.ExecuteScalar();
                        isValid = (rows > 0);
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error validating job ID {jobID}: {ex.Message}");
            }

            return isValid;
        }


        /// <summary>
        /// verify work order belongs to job
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="woNo"></param>
        /// <returns>boolean indicating work order belongs to job</returns>
        private static bool valid_work_order(int jobID, string woNo, string connectionString)
        {
            bool isValid = false;
            string projFinalID = $"{jobID}-{woNo}";
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Validating work order \"{woNo}\" for job {jobID}");
                    string cmdText = "SELECT COUNT(ProjectFinalID) FROM MLG.POL.tblProjectFinal where ProjectFinalID = @projFinalID";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@projFinalID", projFinalID);
                        conn.Open();
                        int rows = (int)cmd.ExecuteScalar();
                        isValid = (rows > 0);
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error validating work order \"{woNo}\" for job {jobID}: {ex.Message}");
            }
            return isValid;
        }


        /// <summary>
        /// generates a random string of characters
        /// </summary>
        /// <param name="length">length of string to generate</param>
        /// <param name="prefix">text to start the random string with</param>
        /// <param name="pool">list of available characters to choose from</param>
        /// <returns></returns>
        private string generate_random_string(int length, string prefix, string pool)
        {
            Random rand = new Random();
            var sb = new StringBuilder(prefix);

            for(var i = sb.Length; i < length; i++)
            {
                var c = pool[rand.Next(0, pool.Length - 1)];
                sb.Append(c);
            }

            return sb.ToString();
        }

        #endregion
    }
}
