using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;


namespace GenerateSchedule
{
    public partial class Form1 : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        public int Scheduled = 0, SameOff = 0, SameShift = 0, LessThan12 = 0, Offvio = 0, Dayes6 = 0;
        private SqlConnection conn = new SqlConnection("Data Source=EMAD-LAP;Initial Catalog=HR900;User ID=sa;Password=ccsystem*360;Connect Timeout=30000");

        //private SqlConnection conn = new SqlConnection("Data Source=172.21.19.222\\ccsystem;Initial Catalog=HR900;User ID=sa;Password=ccsystem*360;Connect Timeout=30000");
        int counter = 0;
        public Form1()
        {

            InitializeComponent();
            comboBox2.SelectedIndex = 3;
            comboBox3.SelectedIndex = 6;
            comboBox4.SelectedIndex = 12;
            cmbH.SelectedIndex = 4;
            var hdr = GetQ(@"SELECT DISTINCT CONVERT(NVARCHAR(555),StartDate) +'-'+ CONVERT(NVARCHAR(555),EndDate) DisplayName,CONVERT(NVARCHAR(555),StartDate) +'|'+ 
CONVERT(NVARCHAR(555), EndDate) Value FROM ScheduleActualData");
            comboBox1.DataSource = hdr;
            comboBox1.DisplayMember = "DisplayName";
            comboBox1.ValueMember = "Value";

        }

        public DataTable GetQ(string sql)
        {
            DataTable dataTable = new DataTable();
            // using (SqlDataAdapter adapter = new SqlDataAdapter($"SELECT top 1 * FROM dbo.Qsestbian WHERE NOT EXISTS (SELECT 1 FROM Notificationlog WHERE Qid=notification.id and UserId='{Environment.UserName}' AND MachineId='{Environment.MachineName}') order by d1", conn))

            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, conn))
            {
                try
                {
                    adapter.Fill(dataTable);
                    return dataTable;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }


        private async void button1_Click(object sender, EventArgs e)
        {

            if (comboBox2.SelectedIndex < 0 || comboBox3.SelectedIndex < 0 || comboBox4.SelectedIndex < 0) return;
            var sDt = DateTime.Parse(comboBox1.SelectedValue.ToString().Split('|')[0]);
            var eDt = DateTime.Parse(comboBox1.SelectedValue.ToString().Split('|')[1]);
            var nSdt = eDt.AddDays(1);
            var nEdt = eDt.AddDays(14);

            Dictionary<string, string> dayNames = new Dictionary<string, string>();

            dayNames["DT01"] = "Sun";
            dayNames["DT02"] = "Mon";
            dayNames["DT03"] = "Tue";
            dayNames["DT04"] = "Wed";
            dayNames["DT05"] = "Thu";
            dayNames["DT06"] = "Fri";
            dayNames["DT07"] = "Sat";
            dayNames["DT11"] = "Sun";
            dayNames["DT12"] = "Mon";
            dayNames["DT13"] = "Tue";
            dayNames["DT14"] = "Wed";
            dayNames["DT15"] = "Thu";
            dayNames["DT16"] = "Fri";
            dayNames["DT17"] = "Sat";


            var dateTable = GetQ($"SELECT * FROM ScheduleActualData "
                                +$"Where StartDate='{comboBox1.SelectedValue.ToString().Split('|')[0]}' And  "+
                                $"sid not in (select sid from srouter where sid <>'' and sid is not null) and "+
                                $" EndDate='{comboBox1.SelectedValue.ToString().Split('|')[1]}'");
            try
            {
                foreach (DataRow row in dateTable.Rows)
                {


                    var EmpId = row["EmpId"].ToString();
                    var SID = row["SID"].ToString();
                    var Center = row["Center"].ToString();
                    var Name = row["Name"].ToString();
                    var offDate = false;
                    await Task.Delay(30);
                    textBox1.Text = EmpId;
                    //textBox2.Text =
                    // System.Threading.Thread.Sleep(100);
                    string offDayesStr = string.Empty;

                    //تحديد اذا كان كل الايام personal 
                    foreach (var columnName in dateTable.Columns.Cast<DataColumn>().Where(c => c.ColumnName.StartsWith("DT")).Select(s => s.ColumnName).ToArray())
                    {
                        if (row[columnName].ToString().Contains("personal") || row[columnName].ToString().Contains("off"))
                        {
                            offDate = true;

                        }
                        else
                        {
                            offDate = false;
                            break;
                        }
                    }

                    List<string> offDayes = new List<string>();
                    List<string> shifts = new List<string>();
                    List<string> shifts1 = new List<string>();
                    if (!offDate)
                    {
                        foreach (var columnName in dateTable.Columns.Cast<DataColumn>().Where(c => c.ColumnName.StartsWith("DT")).Select(s => s.ColumnName).ToArray())
                        {
                            //   if (row[columnName].ToString().Contains("personal") || row[columnName].ToString().ToLower().Contains("off"))???
                            if (row[columnName].ToString().ToLower().Contains("off"))
                            {
                                offDayes.Add(columnName);
                                offDayesStr += dayNames[columnName] + ",";
                            }
                            else
                            {
                                shifts.Add(row[columnName].ToString());
                            }
                            shifts1.Add(row[columnName].ToString());
                        }
                    }
                    //الحصول على الايام التي لاتحتوي على اجازة وتجميعها حسب الشفت من الاكثرتكرار الى الاقل
                    var grp = shifts.Where(s => !s.Equals("personal", StringComparison.OrdinalIgnoreCase)).Select(s => new { Nm = s })
                        .GroupBy(g => g.Nm).Select(x => new { Nm = x.Key, Count = x.Count() }).OrderByDescending(o => o.Count).ToList();
                    var shiftTime = "";
                    if (grp.Count > 0)
                        // assign الوردية الاكثر تكرار
                        shiftTime = grp[0].Nm;
                    else
                        //في حال عدو وجود وردية تصبح personal  (في هذه الحال الوردية عبارة عن personal و off)
                        shiftTime = "personal";
                    shifts1.Reverse();//تحتوي على الوردية كاملة معكوسة
                    shifts.Reverse();//تحتوي على الوردية بدون الراحات
                    var tb = GetQ($"SELECT * FROM SRouter Where isnull(CNTR,'0') <> '*'");

                    foreach (DataRow r in tb.Rows)
                    {
                        counter++;
                        //var shft =  r["shift"].ToString().Split(new[] { '-' })[1].Trim();
                        //var sdt = Convert.ToDateTime(shft);

                        //var ss = shiftTime.Split(new[] { '-' })[1].Trim();
                        //var edt = Convert.ToDateTime(ss);

                        // if (r["shift"].ToString() != shiftTime && !chkShift.Checked)
                        //  {
                        // الراحة المتاحة في فوركاست المحتملة
                        var df = r["dayoff"].ToString(); ;
                        var arr = df.Split(new[] { ',' });
                        // الوردية المتاحة في فوركاست المحتملة مع فصل الوقت 
                        var SRouter = r["shift"].ToString().Split(new[] { '-' });
                        textBox2.Text = shiftTime;
                        textBox3.Text = r["shift"].ToString();
                        var curr = shiftTime.Split(new[] { '-' });//فصل الوقت في الوردية الاكثر تكرارا


                        //حساب الفرق بين بداية الوردية السابقة والوردية الجديدة بالساعات
                        int minD = 0;
                        if (curr[0].ToString() != "personal")
                        {
                            var SRoutert = int.Parse(SRouter[0].Trim().Split(':')[0]);
                            var currt = int.Parse(curr[0].Trim().Split(':')[0]);
                            if (SRoutert > currt && SRoutert != 0 && currt != 0)
                                minD = SRoutert - currt;
                            else if (currt == 0 && SRoutert > 0 && SRoutert <= 12)
                                minD = SRoutert;
                            else if (currt == 0 && SRoutert > 0 && SRoutert > 12)
                                minD = 24 - SRoutert;
                            else if (currt > SRoutert && SRoutert != 0 && currt != 0)
                                minD = currt - SRoutert;
                            else if (SRoutert == 0 && currt > 0 && currt < 12)
                                minD = currt;
                            else if (SRoutert == 0 && currt > 0 && currt >= 12)
                                minD = 24 - currt;
                            else if (SRoutert == 0 && currt == 0)
                                minD = 0;
                        }
                        //   if (!(offDayesStr.Contains(df) || offDayesStr.Contains(arr[0]) || offDayesStr.Contains(arr[1])))
                        //    if (!(offDayesStr.Contains(df)))

                        //if (shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase) )
                        // shiftTime الوردية الاكثر تكرارا
                        // if (shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase) ||
                        //!(r["shift"].ToString().Contains(curr[0].Trim()) ||
                        //r["shift"].ToString().Contains(curr[1].Trim())) ||

                        //((r["shift"].ToString().Contains(curr[0].Trim()) ||
                        //r["shift"].ToString().Contains(curr[1].Trim())) && r["dayoff"].ToString().Contains("Sun"))

                        //)
                        //
                        //
                        //
                        //
                        //
                        //
                        //اذا كانت الوريدة الحالية = personal او 
                        //اذا كان وقت الخروج  من الوردية الجديد لايساوي وقت الدخول في الوردية السابقة اوكان وقت الدخول في الوردية الجديدة هو وقت الخروج من الوردية السابقة 
                        //

                        var hasPersonal = shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase);
                        var fShift = r["shift"].ToString().Split('-')[1];
                        var currentShift = curr[0].Trim();
                        if ((hasPersonal ||!(r["shift"].ToString().Split('-')[1].Trim() == curr[0].Trim()) ||

                            ((r["shift"].ToString().Split('-')[1].Trim() == (curr[0].Trim()) || 
                            r["shift"].ToString().Split('-')[0].Trim() == (curr[1].Trim())) && 
                            r["dayoff"].ToString().Contains("Sun"))

                   ) && !chkShift.Checked)
                        //    if (!(shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase))) continue;
                        //if()

                        //مراجعة

                        //     if (shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase) ||
                        //!(r["shift"].ToString().Split('-')[1].Trim() == (curr[0].Trim()) || r["shift"].ToString().Split('-')[0].Trim() == (curr[1].Trim())) ||

                        //((r["shift"].ToString().Split('-')[1].Trim() == (curr[0].Trim()) || r["shift"].ToString().Split('-')[0].Trim() == (curr[1].Trim())) && r["dayoff"].ToString().Contains("Sun"))

                        //)
                        {
                            //if (!shiftTime.Equals("personal", StringComparison.OrdinalIgnoreCase))
                            //if (shifts[0].ToLower() != "personal" && shifts1[0].ToLower() != "off")
                            //{
                            // if(curr[0].ToString().Trim().Split(':')[0]- int.Parse(SRouter[0].Trim().Split(':')[0])>4)
                            //var sdt1 = Convert.ToDateTime(SRouter[0].Trim());
                            //var edt1 = Convert.ToDateTime(curr[1].Trim());

                            //var h1 = sdt1.Subtract(edt1).Hours;

                            int x = 12;
                            if (curr[0].ToString() != "personal")
                            {
                                var sdt = int.Parse(SRouter[0].Trim().Split(':')[0]);
                                var edt = int.Parse(curr[1].Trim().Split(':')[0]);
                                if (sdt > 0 && sdt < edt)
                                    x = 24 - edt + sdt;
                                else if (edt == 0 && sdt > 0)
                                    x = sdt;
                                else if (sdt > edt)

                                    x = sdt - edt;
                                else if (sdt == edt && sdt < 12)
                                    x = 0;
                                else if (sdt == edt && sdt >= 12)
                                    x = 24;
                                else if (sdt == 0 && edt > 0 && edt < 12)
                                    x = sdt - edt;

                                else if (sdt == 0 && edt > 0 && edt > 12)
                                    x = 24 - edt;
                                else
                                    x = 0;
                            }
                            var h = x;




                            var nmbr = dayNames.FirstOrDefault(f => f.Value == arr[0]).Key;
                            if (nmbr == "DT07") nmbr = dayNames.FirstOrDefault(f => f.Value == arr[1]).Key;
                            var ff = offDayes.OrderByDescending(s => s).ToList();
                            var last = ff[0];

                            var LoffDayes = int.Parse(last.Replace("DT", ""));
                            var lsf = LoffDayes;
                            var newoffDayes = int.Parse(nmbr.Replace("DT", ""));
                            LoffDayes = LoffDayes > 10 ? 17 - LoffDayes : 17 - LoffDayes + 3;

                            //var diff2 = (newoffDayes + 6) - LoffDayes;

                            var diff = LoffDayes + (newoffDayes - 1);
                            //    dataGridView1.Refresh();
                            // if (diff < 3 || diff > 6)
                            if (diff < int.Parse(comboBox2.SelectedItem.ToString()) || diff > int.Parse(comboBox3.SelectedItem.ToString()))
                                continue;
                            if (shiftTime == r["shift"].ToString() && !chkShift.Checked) continue;
                            if (offDayesStr.Substring(0, 7) == df.ToString() && !chkOff.Checked) continue;
                            ////  if (diff > 6)
                            ////     continue;
                            //if (h <= 12 && !df.Contains("Sun") && lsf != 17) //&& Center != "Dammam STC Channels"
                            if (lsf == 17 || df.Contains("Sun"))//&& minD > 6
                            { }
                            else
                            if (h <= int.Parse(comboBox4.SelectedItem.ToString()) && !df.Contains("Sun") && lsf != 17) //&& minD > 6//&& Center != "Dammam STC Channels"

                                continue;

                            //  }
                            if (minD >= int.Parse(cmbH.SelectedItem.ToString()))
                            {
                                Scheduled++;
                                string strCom = $@"Update SRouter set 
                                                    EmpId = '{EmpId}',
                                                    SID = '{SID}',
                                                    Center = '{Center}',
                                                    Name='{Name}',
                                                    CNTR = '*',

                                                    --startdate = '{r["EmpId"]}',
                                                    --enddate = '{r["EmpId"]}',

                                                     PrevShift='{shiftTime}',
                                                    PrevOff='{offDayesStr.Replace(",", "-")}',
                                                    DT01 = '{string.Format("{0}", df.ToString().Contains("Sun") == true ? "Off" : r["shift"])}',
                                                    DT02 ='{string.Format("{0}", df.ToString().Contains("Mon") == true ? "Off" : r["shift"])}',
                                                    DT03 = '{string.Format("{0}", df.ToString().Contains("Tue") == true ? "Off" : r["shift"])}',
                                                    DT04 = '{string.Format("{0}", df.ToString().Contains("Wed") == true ? "Off" : r["shift"])}',
                                                    DT05 = '{string.Format("{0}", df.ToString().Contains("Thu") == true ? "Off" : r["shift"])}',
                                                    DT06 = '{string.Format("{0}", df.ToString().Contains("Fri") == true ? "Off" : r["shift"])}',
                                                    DT07 = '{string.Format("{0}", df.ToString().Contains("Sat") == true ? "Off" : r["shift"])}',
                                                    DT11 = '{string.Format("{0}", df.ToString().Contains("Sun") == true ? "Off" : r["shift"])}',
                                                    DT12 ='{string.Format("{0}", df.ToString().Contains("Mon") == true ? "Off" : r["shift"])}',
                                                    DT13 = '{string.Format("{0}", df.ToString().Contains("Tue") == true ? "Off" : r["shift"])}',
                                                    DT14 = '{string.Format("{0}", df.ToString().Contains("Wed") == true ? "Off" : r["shift"])}',
                                                    DT15 = '{string.Format("{0}", df.ToString().Contains("Thu") == true ? "Off" : r["shift"])}',
                                                    DT16 = '{string.Format("{0}", df.ToString().Contains("Fri") == true ? "Off" : r["shift"])}',
                                                    DT17 = '{string.Format("{0}", df.ToString().Contains("Sat") == true ? "Off" : r["shift"])}'
                                                      Where ID={r["id"]}";
                                if (shiftTime == r["shift"].ToString())
                                { SameShift++; textBox8.Text = SameShift.ToString(); }
                                if (offDayesStr.Substring(0, 7) == df.ToString())
                                { SameOff++; textBox9.Text = SameOff.ToString(); }
                                textBox5.Text = Scheduled.ToString();
                                if (diff == 6)
                                { Dayes6++; textBox11.Text = Dayes6.ToString(); }
                                if (h < 12)
                                { LessThan12++; textBox10.Text = LessThan12.ToString(); }
                                Exec(strCom);
                                break;

                            }
                        }
                        //  }


                        //dataGridView1.;
                    }

                    var dts = GetQ(@"select * from  srouter");
                    var dt = GetQ(@"select * from  ScheduleActualData");
                    textBox4.Text = dt.Rows.Count.ToString();
                    textBox6.Text = ((Scheduled / float.Parse(dts.Rows.Count.ToString()) * 100)).ToString();



                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
            MessageBox.Show(counter.ToString());
            counter = 0;
            //   sRouterBindingSource.
            // dataGridView1.DataSource = "sRouterBindingSource";
            //  dataGridView1.Update();
            //  dataGridView1.Refresh();

        }

        private void Exec(string sql)
        {
            using (SqlCommand command = new SqlCommand(sql, conn))
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                    command.ExecuteNonQuery();
                    command.Connection.Close();

                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void cmbH_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
            button2.Enabled = false;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ClearAll();



        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox2.Items.Add(listBox1.SelectedItem.ToString());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            listBox4.Items.Add(listBox3.SelectedItem.ToString());
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex >= 0)
                listBox2.Items.Remove(listBox2.Items[listBox2.SelectedIndex]);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (listBox4.SelectedIndex >= 0)
                listBox4.Items.Remove(listBox4.Items[listBox4.SelectedIndex]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            button2.Enabled = false;

        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count > 0)
            {

                for (int i = 0; i < listBox4.Items.Count; i++)
                {
                    string[] row = new string[] { listBox2.Items[0].ToString(), listBox4.Items[i].ToString() };
                    dataGridView1.Rows.Add(row);

                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            //Workbook workbook = new Workbook();
            //Worksheet sheet = workbook.Worksheets[0];
            //sheet.InsertDataTable(datatable, true, 1, 1);
            //workbook.SaveToFile("DataTable2Excel.xlsx", ExcelVersion.Version2013);
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Path.GetFileName(saveFileDialog1.FileName);
            }


            var dateTable = GetQ($"select * from srouter");
            var lines = new List<string>();

            string[] columnNames = dateTable.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName).
                                              ToArray();

            var header = string.Join(",", columnNames);
            lines.Add(header);

            var valueLines = dateTable.AsEnumerable()
                               .Select(row => string.Join(",", row.ItemArray));
            lines.AddRange(valueLines);
            if (saveFileDialog1.FileName != "")
                File.WriteAllLines(string.Format("{0}{1}", Path.GetFullPath(saveFileDialog1.FileName), "excel.csv"), lines, Encoding.UTF8);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            // string header = rbHeaderYes.Checked ? "YES" : "NO";
            string header = "YES";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        //con.Close();
                        DataTable Skills = new DataTable();
                        cmd.CommandText = "SELECT distinct skills From [" + sheetName + "]";
                        //con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(Skills);

                        foreach (DataRow row in Skills.Rows)
                        {
                            listBox1.Items.Add(row["skills"].ToString());
                        }
                        //Populate DataGridView.  
                        con.Close();
                        dataGridView2.DataSource = dt;
                    }
                }
            }
        }
        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            // string header = rbHeaderYes.Checked ? "YES" : "NO";
            string header = "YES";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        //con.Close();
                        DataTable Skills = new DataTable();
                        cmd.CommandText = "SELECT distinct skills From [" + sheetName + "]";
                        //con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(Skills);

                        foreach (DataRow row in Skills.Rows)
                        {
                            listBox1.Items.Add(row["skills"].ToString());
                        }
                        //Populate DataGridView.  
                        con.Close();
                        dataGridView2.DataSource = dt;
                    }
                }
            }
        }
        void ClearAll()
        {
            string Query = @"UPDATE dbo.SRouter SET EmpId = NULL, sid = NULL, name = NULL, Center = NULL,
CNTR = NULL, prevshift = NULL, prevoff = NULL,dt01=null,dt02=null,dt03=null,dt04=null,
dt05=null,dt06=null,dt07=null,dt11=null,dt12=null,dt13=null,dt14=null,dt15=null,dt16=null,dt17=null";
            Exec(Query);
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox4.Items.Clear();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            Scheduled = 0; SameOff = 0; SameShift = 0; LessThan12 = 0; Offvio = 0; Dayes6 = 0;

        }
        private void accordionControl1_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
