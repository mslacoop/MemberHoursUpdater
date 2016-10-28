using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using System.Net;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Download;

using LumenWorks.Framework.IO.Csv;

namespace COOP
{
    public partial class Form1 : Form
    {

        public class WebClientEx : WebClient
        {
            public WebClientEx(CookieContainer container)
            {
                this.container = container;
            }

            private readonly CookieContainer container = new CookieContainer();

            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest r = base.GetWebRequest(address);
                var request = r as HttpWebRequest;
                if (request != null)
                {
                    request.CookieContainer = container;
                }
                return r;
            }

            protected override WebResponse GetWebResponse(WebRequest request, IAsyncResult result)
            {
                WebResponse response = base.GetWebResponse(request, result);
                ReadCookies(response);
                return response;
            }

            protected override WebResponse GetWebResponse(WebRequest request)
            {
                WebResponse response = base.GetWebResponse(request);
                ReadCookies(response);
                return response;
            }

            private void ReadCookies(WebResponse r)
            {
                var response = r as HttpWebResponse;
                if (response != null)
                {
                    CookieCollection cookies = response.Cookies;
                    container.Add(cookies);
                }
            }
        }


        /// <summary>
        /// Store the Application object we can use in the member functions.
        /// </summary>
        int iFirstNameIndex;
        int iLastNameIndex;
        int iStatusIndex;
        int iIDIndex;
        
        string sSecrect;

        public struct Member
        {
            private int id;
            private string firstname;
            private string lastname;
            private int hours;

            public int iID
            {
                get { return id; }
                set { id = value; }
            }
            public int iHours
            {
                get { return hours; }
                set { hours = value; }
            }
            public string sFirstName
            {
                get { return firstname; }
                set { firstname = value; }
            }
            public string sLastName
            {
                get { return lastname; }
                set { lastname = value; }
            }

        }
        public struct Product
        {
            private string upc;
            private string desc;

            public string sUPC
            {
                get { return upc; }
                set { upc = value; }
            }
            public string sDesc
            {
                get { return desc; }
                set { desc = value; }
            }

        }
        List<Member> lstGoodList;
        List<Member> lstBadList;


        List<Product> lstProdList;

        static string[] Scopes = { DriveService.Scope.DriveReadonly };
        static string ApplicationName = "COOP-WinForm-App1";

        private MySqlConnection SQLCon;

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.progressDownload.Visible = false;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            button1.Enabled = false;
            sSecrect = "";
            button3.Visible = false;
            lstProdList = new List<Product>();

            this.KeyPreview = true;

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = String.Format("IS4C DB Updater {0}", version);

            //this is just the default shit
            if (coop.Default.DBConnection.Length < 5)
                coop.Default.DBConnection = "Server=192.168.0.100;Port=3306;Database=is4c_op;Uid=backend;Pwd=is4cbackend;default command timeout=7;";
            if (coop.Default.Lane1DBConnction.Length < 5)
                coop.Default.Lane1DBConnction = "Server=192.168.0.101;Port=3306;Database=opdata;Uid=backend;Pwd=is4cbackend;default command timeout=7;";
            if (coop.Default.Lane2DBConnction.Length < 5)
                coop.Default.Lane2DBConnction = "Server=192.168.0.102;Port=3306;Database=opdata;Uid=backend;Pwd=is4cbackend;default command timeout=7;";


            //sample connection string
            //myConnectionString="Server=myServerAddress;Port=1234;Database=testDB;Uid=root;Pwd=abc123;
            //dfault Mysql port is 3306

            tbConnectionString.Text = coop.Default.DBConnection;
            tbLane1ConnectionString.Text = coop.Default.Lane1DBConnction;
            tbLane2ConnectionString.Text = coop.Default.Lane2DBConnction;

            SQLCon = new MySqlConnection(tbConnectionString.Text);

            //tbFilePath.Text = coop.Default.FilePath;

            //if (File.Exists(tbFilePath.Text))
            //    ProcessFile();

            try
            {
                MessageBox.Show("Connecting to the Backend, this might take some time.\n\nGive it a minute before you start double clicking all over the place.", "DB Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;
                // UNCOMMENT THIS WHEN YOU WANT TO TESt IN THE COOP
                SQLCon.Open();
                label4.Text = "DB Connection: GOOD";
                //button1.Enabled = true;
            }
            catch(Exception e)
            {
                // Set cursor back
                Cursor.Current = Cursors.Default;
                label4.Text = "DB Connection: FAILED!!!!!";
                button1.Enabled = false;
                button2.Enabled = false;
                this.AllowDrop = false;
                MessageBox.Show("There was a problem connecting to the backend DB. \n\nYOU SHALL NOT PASS!!!\n\nDon't know what to tell ya....Jiggle the handle maybe?\n\nCall my creator please." + Environment.NewLine + Environment.NewLine + e.Message, "DB Connection Issues", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                if (System.Windows.Forms.Application.MessageLoop)
                {
                    // Use this since we are a WinForms app
                    //System.Windows.Forms.Application.Exit();   //sjg-we dont want to exit cause we might need to update connection string
                }
                else
                {
                    // Use this since we are a console app
                    //System.Environment.Exit(1);   //sjg-we dont want to exit cause we might need to update connection string
                }
            }
            finally
            {
                //SQLCon.Close();
            }


        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Link;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            tbFilePath.Text = files[0];
            //coop.Default.FilePath = tbFilePath.Text;
            //coop.Default.Save();

            if (Path.GetExtension(files[0]) == ".cvs")
                ProcessFileCvs();
            else if (Path.GetExtension(files[0]) == ".xlsx")
                ProcessFileExcel();
            else
                MessageBox.Show("Unknow File Type. You can drop '.xlsx' or '.cvs' files here","Unknow File Extension",MessageBoxButtons.OK,MessageBoxIcon.Error);


        }

        //this will clean up all the open Interop bullshit that is Excel.
        void CleanStateExit(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();


            Marshal.FinalReleaseComObject(xlRange);
            Marshal.FinalReleaseComObject(xlWorksheet);

            xlWorkbook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
        }

        //read in the excel spreadsheet and try to figure out what is what
        private void ProcessFileExcel()
        {
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;

            //init some things before we go looking for the columns that hold our information
            button1.Enabled = false;
            iFirstNameIndex = -1;
            iLastNameIndex = -1;
            iStatusIndex = -1;
            iIDIndex = -1;

            lstBadList = new List<Member>();
            lstGoodList = new List<Member>();
            lbBadMembers.Items.Clear();
            lbGoodMembers.Items.Clear();

            label2.Text = "Members in Good Standing";
            label3.Text = "Members NOT in Good Standing";

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            try
            {

                //https://docs.google.com/spreadsheets/d/1rxY0Oro8N9Ehi2aVRlh0DjsH7TPXWAJivt8TwOJfW4g/export?exportFormat=xls
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(@tbFilePath.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;


                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;


                //find out which columns how the info we are interested in. First and Last Name and the current status value
                for (int j = 1; j <= colCount; j++)
                {

                    string temp = System.Convert.ToString((xlRange.Cells[1, j] as Microsoft.Office.Interop.Excel.Range).Value2);
                    temp = temp.ToLower();
                    if (temp != null)
                    {
                        if (temp.CompareTo("first name") == 0)
                            iFirstNameIndex = j;
                        if (temp.CompareTo("overall status") == 0)
                            iStatusIndex = j;
                        if (temp.CompareTo("0") == 0)
                            iLastNameIndex = j;
                        if (temp.CompareTo("last name") == 0)
                            iLastNameIndex = j;
                        if (temp.CompareTo("id") == 0)
                            iIDIndex = j;
                        if (temp.CompareTo("owner number") == 0)
                            iIDIndex = j;
                    }
                }

                if (iLastNameIndex == -1)
                {
                    MessageBox.Show("Looks like we have a problem.\nI could not find the column for Last Name in the spreadsheet.\nMake sure the spreadsheet has a column called 'Last Name' OR the last name column is called '0'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor.Current = Cursors.Default;
                    CleanStateExit(xlApp, xlWorkbook, xlWorksheet, xlRange);
                    return;
                }
                if (iFirstNameIndex == -1)
                {
                    MessageBox.Show("Looks like we have a problem.\nI could not find the column for First Name in the spreadsheet.\nMake sure the spreadsheet has a column called 'First Name'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor.Current = Cursors.Default;
                    CleanStateExit(xlApp, xlWorkbook, xlWorksheet, xlRange);
                    return;
                }
                if (iStatusIndex == -1)
                {
                    MessageBox.Show("Looks like we have a problem.\nI could not find the column for Overall Status in the spreadsheet.\nMake sure the spreadsheet has a column called 'OVERALL STATUS'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor.Current = Cursors.Default;
                    CleanStateExit(xlApp, xlWorkbook, xlWorksheet, xlRange);
                    return;
                }

                if (iIDIndex == -1)
                {
                    MessageBox.Show("Looks like we have a problem.\nI could not find the column for Member ID number in the spreadsheet.\nMake sure the spreadsheet has a column called 'Member ID' OR just 'ID'\nand it has each member's ID number in it.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor.Current = Cursors.Default;
                    CleanStateExit(xlApp, xlWorkbook, xlWorksheet, xlRange);
                    return;
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    double temp = 0;
                    if ((xlRange.Cells[i, iStatusIndex] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                    {
                        temp = (xlRange.Cells[i, iStatusIndex] as Microsoft.Office.Interop.Excel.Range).Value2;
                    }
                    else
                    {
                        temp = -999;
                    }
                    //build the name
                    string name = (xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2 + ", " +
                        (xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2;



                    if (name.Length > 3)
                    {
                        Member m = new Member();

                        Type t = (xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2.GetType();
                        if (t.Equals(typeof(double)))
                            m.sFirstName = ((int)(xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2).ToString();
                        else
                            m.sFirstName = (xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2;

                        t = (xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2.GetType();
                        if (t.Equals(typeof(double)))
                            m.sLastName = ((int)(xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2).ToString();
                        else
                            m.sLastName = (xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2;

                        if ((xlRange.Cells[i, iIDIndex] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                        {
                            t = (xlRange.Cells[i, iIDIndex] as Microsoft.Office.Interop.Excel.Range).Value2.GetType();
                            if (t.Equals(typeof(double)))
                                m.iID = (int)(xlRange.Cells[i, iIDIndex] as Microsoft.Office.Interop.Excel.Range).Value2;
                            else
                            {
                                MessageBox.Show("Member Number is not set to a number.\n\n" + m.sLastName + "," + m.sFirstName);
                                m.iID = -1;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Member Number is not set to a number.\n\n" + m.sLastName + "," + m.sFirstName);
                            m.iID = -1;
                        }

                        //m.sFirstName = (xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2;
                        //m.sLastName = (xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2;
                        //m.iID = (int)(xlRange.Cells[i, iIDIndex] as Microsoft.Office.Interop.Excel.Range).Value2;
                        //m.iID = 211;
                        m.iHours = (int)temp;

                        if (temp > -6)
                        {
                            lbGoodMembers.Items.Add(name);
                            label2.Text = "Members in Good Standing (" + lbGoodMembers.Items.Count.ToString() + ")";

                            lstGoodList.Add(m);

                        }
                        else
                        {
                            lbBadMembers.Items.Add(name);
                            label3.Text = "Members NOT in Good Standing (" + lbBadMembers.Items.Count.ToString() + ")";
                            lstBadList.Add(m);
                        }
                        //Console.Write((xlRange.Cells[i, iFirstNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2 + " " + (xlRange.Cells[i, iLastNameIndex] as Microsoft.Office.Interop.Excel.Range).Value2 + "=" + temp + Environment.NewLine);
                    }
                }//end forloop
                CleanStateExit(xlApp, xlWorkbook, xlWorksheet, xlRange);
                button1.Enabled = true;
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);

                if(e.Message.Contains("Class not registered"))
                    MessageBox.Show("Well that didn't work. Do you have Excel installed?\nIn order for me to read an Excel Spreadsheet you need to have Excel installed.\nTry using downloading as a CVS instead.","Error Reading the File",MessageBoxButtons.OK,MessageBoxIcon.Error);
                else
                    MessageBox.Show("{0} Exception caught." + Environment.NewLine + e.Message);


            }
            finally
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;

            }
        }

        //read in the excel spreadsheet and try to figure out what is what
        private void ProcessFileCvs()
        {
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;

            //init some things before we go looking for the columns that hold our information
            button1.Enabled = false;
            iFirstNameIndex = -1;
            iLastNameIndex = -1;
            iStatusIndex = -1;
            iIDIndex = -1;

            lstBadList = new List<Member>();
            lstGoodList = new List<Member>();
            lbBadMembers.Items.Clear();
            lbGoodMembers.Items.Clear();

            label2.Text = "Members in Good Standing";
            label3.Text = "Members NOT in Good Standing";


            try
            {
                int colCount = 0;

                // open the file "data.csv" which is a CSV file with headers
                using (CsvReader csv =
                       new CsvReader(new StreamReader(@tbFilePath.Text), true))
                {
                    colCount = csv.FieldCount;

                    string[] headers = csv.GetFieldHeaders();
                    //find out which columns how the info we are interested in. First and Last Name and the current status value
                    for (int j = 0; j <= colCount-1; j++)
                    {

                        string tempheader = headers[j];
                        tempheader = tempheader.ToLower();
                        if (tempheader != null)
                        {
                            if (tempheader.CompareTo("first name") == 0)
                                iFirstNameIndex = j;
                            if (tempheader.CompareTo("overall status") == 0)
                                iStatusIndex = j;
                            if (tempheader.CompareTo("0") == 0)
                                iLastNameIndex = j;
                            if (tempheader.CompareTo("last name") == 0)
                                iLastNameIndex = j;
                            if (tempheader.CompareTo("id") == 0)
                                iIDIndex = j;
                            if (tempheader.CompareTo("owner number") == 0)
                                iIDIndex = j;
                        }
                    }

                    if (iLastNameIndex == -1)
                    {
                        MessageBox.Show("Looks like we have a problem.\nI could not find the column for Last Name in the spreadsheet.\nMake sure the spreadsheet has a column called 'Last Name' OR the last name column is called '0'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor.Current = Cursors.Default;
                        return;
                    }
                    if (iFirstNameIndex == -1)
                    {
                        MessageBox.Show("Looks like we have a problem.\nI could not find the column for First Name in the spreadsheet.\nMake sure the spreadsheet has a column called 'First Name'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor.Current = Cursors.Default;
                        return;
                    }
                    if (iStatusIndex == -1)
                    {
                        MessageBox.Show("Looks like we have a problem.\nI could not find the column for Overall Status in the spreadsheet.\nMake sure the spreadsheet has a column called 'OVERALL STATUS'", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor.Current = Cursors.Default;
                        return;
                    }

                    if (iIDIndex == -1)
                    {
                        MessageBox.Show("Looks like we have a problem.\nI could not find the column for Member ID number in the spreadsheet.\nMake sure the spreadsheet has a column called 'Member ID' OR just 'ID'\nand it has each member's ID number in it.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor.Current = Cursors.Default;
                        return;
                    }



                    while (csv.ReadNextRecord())
                    {
                        double temp = 0;
                        if (csv[iStatusIndex] != null)
                        {
                            double.TryParse(csv[iStatusIndex],out temp);
                        }
                        else
                        {
                            temp = -999;
                        }
                        //build the name
                        string name = csv[iLastNameIndex] + ", " + csv[iFirstNameIndex];

                        if (name.Length > 3)
                        {
                            Member m = new Member();

                            m.sFirstName = csv[iFirstNameIndex];
                            m.sLastName = csv[iLastNameIndex];

                            if (csv[iIDIndex] != null)
                            {
                                if(csv[iIDIndex].Length>0)
                                {
                                    int y = 0;
                                    if(Int32.TryParse(csv[iIDIndex],out y))
                                    {
                                        m.iID = y;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Owner Number is not set to a number.\n\n" + m.sLastName + "," + m.sFirstName);
                                        m.iID = -1;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Owner Number is not set to a number.\n\n" + m.sLastName + "," + m.sFirstName);
                                    m.iID = -1;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Owner Number is not set to a number.\n\n" + m.sLastName + "," + m.sFirstName);
                                m.iID = -1;
                            }

                            m.iHours = (int)temp;

                            if (temp > -6)
                            {
                                lbGoodMembers.Items.Add(name);
                                label2.Text = "Members in Good Standing (" + lbGoodMembers.Items.Count.ToString() + ")";

                                lstGoodList.Add(m);

                            }
                            else
                            {
                                lbBadMembers.Items.Add(name);
                                label3.Text = "Members NOT in Good Standing (" + lbBadMembers.Items.Count.ToString() + ")";
                                lstBadList.Add(m);
                            }
                        }//end name length if
                    }//end while
                }//end of Using statement
                button1.Enabled = true;

            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                MessageBox.Show("{0} Exception caught." + Environment.NewLine + e.Message);

            }
            finally
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;

            }
        }

        //open file button press
        private void button2_Click(object sender, EventArgs e)
        {
            iFirstNameIndex=-1;
            iLastNameIndex = -1;
            iStatusIndex = -1;
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Excel Files (.xlxs)|*.xlsx";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                tbFilePath.Text = openFileDialog1.FileName;
                //coop.Default.FilePath = tbFilePath.Text;
                //coop.Default.Save();

                if (Path.GetExtension(openFileDialog1.FileName) == ".cvs")
                    ProcessFileCvs();
                else if (Path.GetExtension(openFileDialog1.FileName) == ".xlsx")
                    ProcessFileExcel();
                else
                    MessageBox.Show("Unknow File Type. You can drop '.xlsx' or '.cvs' files here", "Unknow File Extension", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        //sssshhhhhhh.....this is the secrect key press to open the connection string
        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Z)
                sSecrect = "";
            else if (e.KeyChar == 122)
                sSecrect = "";
            else
                sSecrect += e.KeyChar;

            if (sSecrect.CompareTo("guerreri") == 0)
            {
               
                tbConnectionString.Visible = true;
                tbLane1ConnectionString.Visible = true;
                tbLane2ConnectionString.Visible = true;
                button3.Visible = true;
            }

        }

        //update the db button has been pressed
        private void button1_Click(object sender, EventArgs e)
        {

            //on is4c check the View 'volunteerDiscounts'.  It might take care of everything.<----NOPE
            //set ALL members to staff of 6.
            //SSI is the number of hours worked
            //      this number might show up on the screen (blueline) with the number of worked hours....NICE
            //'MemDiscountLimit' or the session var 'discountcap' in the lane code--->doesnt look like this would be a problem but keep an eye on it


            //will have to push the hours worked (SSI) out to the db, AND change wmdiscount() in PrehKeys.php from .05 to .15 for staff = 6;

            int count_good = 0;
            int count_bad = 0;
            int errors = 0;

            //Lets check the DB connection 1st
            if(SQLCon != null && SQLCon.State != ConnectionState.Open)
            {
                label4.BackColor = label4.BackColor == Color.Red ? Color.Yellow : Color.Red;
                return;
            }
            //There is a bug where members that are taken off the Attendence DB are still in the lane's db since we don't 
            //'zero' out everone before going in and updating hours.
            //Code below will mark everone except visitor that is a working member (staff=6) to non member and put hours at -99
            string sql2 = "UPDATE members SET staff=0,SSI=-99 WHERE NOT CardNo = 1;";

            try
            {
                //create a command object that will be used to call the stored procedure
                using (MySqlCommand cmd = new MySqlCommand(sql2, SQLCon))
                {
                    cmd.CommandType = CommandType.Text;
                    //open the connection, run the stored procedure and close the connection
                    //SQLCon.Open();
                    int rows = cmd.ExecuteNonQuery();

                    if (rows < 1)
                    {
                        MessageBox.Show("The Reset SQL for this program failed./nPlease contact my creator to fix this.");
                    }



                }
            }
            catch (Exception ex)
            {

                //there be errors 
                MessageBox.Show("There was a problem but we will keep going." + Environment.NewLine + Environment.NewLine + ex.Message, "Update Error 1", MessageBoxButtons.OK);

                errors++;
            }



            //lets loop thru all the lists and mark good and bad members one at a time.....
            //this might take forever but it will be easier to recover from errors
            foreach (Member m in lstGoodList)
            {
                if (m.iID > 0)
                {
                    //we may  not need to update anything but SSI. I would still want to mark everyone on these lists as active and staff of 6 just to make sure
                    string sql = "UPDATE members SET staff=6,SSI=" + m.iHours.ToString() + " WHERE CardNo= " + m.iID.ToString() + ";";

                    try
                    {
                        //create a command object that will be used to call the stored procedure
                        using (MySqlCommand cmd = new MySqlCommand(sql, SQLCon))
                        {
                            cmd.CommandType = CommandType.Text;
                            //open the connection, run the stored procedure and close the connection
                            //SQLCon.Open();
                            int rows = cmd.ExecuteNonQuery();

                            if (rows > 1)
                            {
                                MessageBox.Show("Member number " + m.iID.ToString()+ " updated more then one person, " + Environment.NewLine + "please check the attendence spreadsheet for errros.");
                            }
                            else if (rows < 1)
                            {
                                MessageBox.Show("Member  " + m.sFirstName + " " + m.sLastName + " was not updated, " + Environment.NewLine + "please check the attendence spreadsheet for errros.");
                            }
                            else if (rows == 1)
                                count_good++;



                        }
                    }
                    catch (Exception ex)
                    {

                        //there be errors 
                        MessageBox.Show("There was a problem but we will keep going." + Environment.NewLine + Environment.NewLine + ex.Message, "Update Error 2", MessageBoxButtons.OK);
                        errors++;
                    }
                    finally
                    {
                        //SQLCon.Close();
                    }
                }
                else
                    MessageBox.Show("Member ID for " + m.sFirstName + " " + m.sLastName + " looks like it isn't a number, " + Environment.NewLine + "please check the attendence spreadsheet.");

            }

            foreach (Member m in lstBadList)
            {
                if (m.iID > 0)
                {
                    //need to check but it looks like is4c checks 2 things for Working Member Discoount.....SSI (hours worked) and isStaff==6
                    //aslong as hours is updated IS$C will do all the work. I'm keeping this BadList update seperate for now just in case the SQL statement needs to change
                    string sql = "UPDATE members SET staff=6,SSI=" + m.iHours.ToString() + " WHERE CardNo= " + m.iID.ToString() + ";";

                    try
                    {
                        //create a command object that will be used to call the stored procedure
                        using (MySqlCommand cmd = new MySqlCommand(sql, SQLCon))
                        {
                            cmd.CommandType = CommandType.Text;
                            //open the connection, run the stored procedure and close the connection
                            //SQLCon.Open();
                            int rows = cmd.ExecuteNonQuery();

                            if (rows > 1)
                            {
                                MessageBox.Show("Member number " + m.iID.ToString() + " updated more then one person, " + Environment.NewLine + "please check the attendence spreadsheet for errros.");
                            }
                            else if (rows < 1)
                            {
                                MessageBox.Show("Member  " + m.sFirstName + " " + m.sLastName + " was not updated, " + Environment.NewLine + "please check the attendence spreadsheet for errros.");
                            }
                            else if (rows == 1)
                                count_bad++;


                        }
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show("There was a problem but we will keep going." + Environment.NewLine + Environment.NewLine + ex.Message, "Update Error 3", MessageBoxButtons.OK);
                        errors++;
                    }
                    finally
                    {
                        //SQLCon.Close();
                    }

                }
                else
                    MessageBox.Show("Member ID for " + m.sFirstName + " " + m.sLastName + " looks like it isn't a number, " + Environment.NewLine + "please check the attendence spreadsheet.");
            }
            if (errors==0)
                MessageBox.Show("You have updated the Member DB\n"+count_good.ToString()+"--good\n"+count_bad.ToString()+"--not good");
            else
                MessageBox.Show("You have updated the Member DB with some errors.\n" + count_good.ToString() + "--good" + Environment.NewLine + count_bad.ToString() + "--not good");

        }


        //save the new connection string
        private void button3_Click(object sender, EventArgs e)
        {
            coop.Default.DBConnection = tbConnectionString.Text;
            coop.Default.Lane1DBConnction = tbLane1ConnectionString.Text;
            coop.Default.Lane2DBConnction = tbLane2ConnectionString.Text;
            coop.Default.Save();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SQLCon.State == ConnectionState.Open)
                SQLCon.Close();
        }


        private void Form1_Shown(object sender, EventArgs e)
        {

            if (MessageBox.Show("Would you like me to auto download the spreadsheet?", "AutoDownload?", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {

                    UserCredential credential;

                    using (var stream =
                        new FileStream(@"client_secret.json", FileMode.Open, FileAccess.Read))
                    {
                        string credPath = System.Environment.GetFolderPath(
                            System.Environment.SpecialFolder.Personal);
                        credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }

                    // Create Drive API service.
                    var service = new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    //// Define parameters of request.
                    //FilesResource.ListRequest listRequest = service.Files.List();
                    //listRequest.PageSize = 10;
                    //listRequest.Fields = "nextPageToken, files(id, name)";

                    //// List files.
                    //IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                    //    .Files;
                    //Console.WriteLine("Files:");
                    //if (files != null && files.Count > 0)
                    //{
                    //    foreach (var file in files)
                    //    {
                    //        Console.WriteLine("{0} ({1})", file.Name, file.Id);
                    //    }
                    //}
                    //else
                    //{
                    //    Console.WriteLine("No files found.");
                    //}





                    var fileId = "1Ee1ZlTeGl3-9hWYr9FMRxB_Ya4Of-njJN_xnnkgmDcM";
                    //var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    var request = service.Files.Export(fileId, "text/csv");
                    
                    var stream2 = new System.IO.MemoryStream();
                    request.MediaDownloader.ChunkSize = 10;


                    bool download_good = true;

                    // Add a handler which will be notified on progress changes.
                    // It will notify on each chunk download and when the
                    // download is completed or failed.
                    request.MediaDownloader.ProgressChanged +=
                            (IDownloadProgress progress) =>
                            {
                                switch (progress.Status)
                                {
                                    case DownloadStatus.Downloading:
                                        {
                                            Console.WriteLine(progress.BytesDownloaded);
                                            break;
                                        }
                                    case DownloadStatus.Completed:
                                        {
                                            Console.WriteLine("Download complete.");
                                            break;
                                        }
                                    case DownloadStatus.Failed:
                                        {
                                            Console.WriteLine("Download failed.");
                                            download_good = false;
                                            MessageBox.Show(progress.Exception.Message,"Error Downloading file",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                                            break;
                                        }
                                }
                            };
                    request.Download(stream2);

                    System.IO.File.Delete(@"temp.cvs");

                    if (download_good)
                    {
                        FileStream file2 = new FileStream("temp.cvs", FileMode.Create, FileAccess.Write);
                        stream2.WriteTo(file2);
                        file2.Close();
                        stream2.Close();

                        tbFilePath.Text = Application.StartupPath + @"\temp.cvs";

                        ProcessFileCvs();

                    }
                    else
                        MessageBox.Show("Well that didnt work. Download it yoursef.\nDownload/Export as an Excel Spreadsheet OR CVS and drag it into this app.", "Error Downloading file", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Well that didnt work. Download it yoursef."+Environment.NewLine+ex.Message, "Error Downloading file", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                finally
                {

                }
                return;
            }
        }

        //check for products that are on the lanes but on in the backend anymore
        private void checkForOrphanUPCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Lets check the DB connection 1st
            if (SQLCon != null && SQLCon.State != ConnectionState.Open)
            {
                label4.BackColor = label4.BackColor == Color.Red ? Color.Yellow : Color.Red;
                return;
            }
            

            //download the Product list from the backend
            using (MySqlCommand cmd = new MySqlCommand("SELECT products.upc, products.description From is4c_op.products WHERE products.inUse=1", SQLCon))
            {
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        //we have a problem, let the user know
                    }
                    while (reader.Read())
                    {
                        Product p = new Product();

                        p.sUPC = reader.GetString("upc");
                        p.sDesc = reader.GetString("description");

                        lstProdList.Add(p);
                    }

                }
            }

            //lane 1
            try
            {
                List<Product> lstLane1ProdList = new List<Product>();
                // SQLCon = new MySqlConnection(tbConnectionString.Text);
                //we have the products from the backend, now check the lanes
                MySqlConnection conLane1 = new MySqlConnection(tbLane1ConnectionString.Text);
                conLane1.Open();
                if (conLane1 != null && conLane1.State != ConnectionState.Open)
                {
                    label4.BackColor = label4.BackColor == Color.Red ? Color.Yellow : Color.Red;
                    return;
                }
                //read in the products from Lane1 and then go from there
                using (MySqlCommand cmd =new MySqlCommand("SELECT products.upc, products.description From opdata.products WHERE products.inUse=1", conLane1))
                {
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read())
                        {
                            //we have a problem, let the user know
                        }
                        while (reader.Read())
                        {
                            Product p = new Product();

                            p.sUPC = reader.GetString("upc");
                            p.sDesc = reader.GetString("description");

                            lstLane1ProdList.Add(p);
                        }
                        //everything is read in.....lets check it out
                        //this might be slow but it will work
                        List<Product> lstOrphanProd= new List<Product>();
                        foreach (Product p in lstLane1ProdList)
                        {
                            if(lstProdList.Contains(p)==false)
                            {
                                //p is not in the Product List from the backend
                                lstOrphanProd.Add(p);
                            }

                        }
                        if(lstOrphanProd.Count>0)
                        {
                            string lst = "";

                            foreach(Product p in lstOrphanProd)
                            {
                                lst += p.sUPC + "-" + p.sDesc+"\n";
                            }
                            DialogResult r = MessageBox.Show("These items where found, click Yes to delete:" + Environment.NewLine  + lst, "Lane 1", MessageBoxButtons.YesNo);
                            if(r == DialogResult.Yes)
                            {
                                //do the delete
                                string sql;
                                foreach(Product pq in lstOrphanProd)
                                {
                                    //sql = "DELETE FROM products where upc = " + pq.sUPC;
                                    sql = "UPDATE products SET inUse=0 WHERE upc = " + pq.sUPC;
                                    MySqlCommand cmdDel = new MySqlCommand(sql, conLane1);
                                    int rows = cmdDel.ExecuteNonQuery();
                                    if(rows==0)
                                    {
                                        MessageBox.Show("For some reason the delete didn't work for:" + Environment.NewLine  + pq.sDesc, "Delete Error");
                                    }
                                }
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem connecting to the Lane 1. " + Environment.NewLine + Environment.NewLine + ex.Message, "DB Connection Issues", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

            //lane 2
            try
            {
                List<Product> lstLane2ProdList = new List<Product>();
                // SQLCon = new MySqlConnection(tbConnectionString.Text);
                //we have the products from the backend, now check the lanes
                MySqlConnection conLane2 = new MySqlConnection(tbLane2ConnectionString.Text);
                conLane2.Open();
                if (conLane2 != null && conLane2.State != ConnectionState.Open)
                {
                    label4.BackColor = label4.BackColor == Color.Red ? Color.Yellow : Color.Red;
                    return;
                }
                using (MySqlCommand cmd = new MySqlCommand("SELECT products.upc, products.description From opdata.products WHERE products.inUse=1", SQLCon))
                {
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read())
                        {
                            //we have a problem, let the user know
                        }
                        while (reader.Read())
                        {
                            Product p = new Product();

                            p.sUPC = reader.GetString("upc");
                            p.sDesc = reader.GetString("description");

                            lstLane2ProdList.Add(p);
                        }
                        //everything is read in.....lets check it out
                        //this might be slow but it will work
                        List<Product> lstOrphanProd = new List<Product>();
                        foreach (Product p in lstLane2ProdList)
                        {
                            if (lstProdList.Contains(p) == false)
                            {
                                //p is not in the Product List from the backend
                                lstOrphanProd.Add(p);
                            }

                        }
                        if (lstOrphanProd.Count > 0)
                        {
                            string lst = "";

                            foreach (Product p in lstOrphanProd)
                            {
                                lst += p.sUPC + "-" + p.sDesc + "\n";
                            }
                            DialogResult r = MessageBox.Show("These items where found, click Yes to delete:\n" + lst, "Lane 2", MessageBoxButtons.YesNo);
                            if (r == DialogResult.Yes)
                            {
                                //do the delete
                                string sql;
                                foreach (Product pq in lstOrphanProd)
                                {
                                    //sql = "DELETE FROM products where upc = " + pq.sUPC;
                                    sql = "UPDATE products SET inUse=0 WHERE upc = " + pq.sUPC;

                                    MySqlCommand cmdDel = new MySqlCommand(sql, conLane2);
                                    int rows = cmdDel.ExecuteNonQuery();
                                    if (rows == 0)
                                    {
                                        MessageBox.Show("For some reason the delete didn't work for:" + Environment.NewLine  + pq.sDesc, "Delete Error");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem connecting to the Lane 2. " + Environment.NewLine + Environment.NewLine + ex.Message, "DB Connection Issues", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }
    }
}
