﻿using System;
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

        List<Member> lstGoodList;
        List<Member> lstBadList;

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

            this.KeyPreview = true;

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = String.Format("IS4C DB Updater {0}", version);

            //this is just the default shit
            if (coop.Default.DBConnection.Length < 5)
                coop.Default.DBConnection = "Server=192.168.0.100;Port=3306;Database=is4c_op;Uid=backend;Pwd=is4cbackend;default command timeout=7;";

            //sample connection string
            //myConnectionString="Server=myServerAddress;Port=1234;Database=testDB;Uid=root;Pwd=abc123;
            //dfault Mysql port is 3306

            tbConnectionString.Text = coop.Default.DBConnection;

            SQLCon = new MySqlConnection(tbConnectionString.Text);

            //tbFilePath.Text = coop.Default.FilePath;

            //if (File.Exists(tbFilePath.Text))
            //    ProcessFile();

            try
            {
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;
                // UNCOMMENT THIS WHEN YOU WANT TO TESt IN THE COOP
                //   SQLCon.Open();
                label4.Text = "DB Connection: GOOD";
                //button1.Enabled = true;
            }
            catch
            {
                // Set cursor back
                Cursor.Current = Cursors.Default;
                label4.Text = "DB Connection: FAILED!!!!!";
                button1.Enabled = false;
                button2.Enabled = false;
                this.AllowDrop = false;
                MessageBox.Show("There was a problem connecting to the backend DB. \n\nYOU SHALL NOT PASS!!!\n\nDon't know what to tell ya....Jiggle the handle maybe?", "DB Connection Issues", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
            ProcessFile();
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
        private void ProcessFile()
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
                    if (temp != null)
                    {
                        if (temp.CompareTo("First Name") == 0)
                            iFirstNameIndex = j;
                        if (temp.CompareTo("OVERALL STATUS") == 0)
                            iStatusIndex = j;
                        if (temp.CompareTo("0") == 0)
                            iLastNameIndex = j;
                        if (temp.CompareTo("Last Name") == 0)
                            iLastNameIndex = j;
                        if (temp.CompareTo("ID") == 0)
                            iIDIndex = j;
                        if (temp.CompareTo("Owner Number") == 0)
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
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);

            }
            finally
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                button1.Enabled = true;

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

                ProcessFile();
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


            int errors = 0;
            //lets loop thru all the lists and mark good and bad members one at a time.....
            //this might take forever but it will be easier to recover from errors
            foreach (Member m in lstGoodList)
            {
                if (m.iID > 0)
                {

                    //we may  not need to update anything but SSI. I would still want to mark everyone on these lists as active and staff of 6 just to make sure
                    string sql = "UPDATE members SET Active=1,isStaff=6,SSI="+m.iHours.ToString()+" WHERE CardNo= " + m.iID.ToString() + ";";

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
                                //updated more then one row
                            }
                            else if (rows < 1)
                            {
                                //less then one row
                            }


                        }
                    }
                    catch (Exception ex)
                    {

                        //there be errors 
                        errors++;
                    }
                    finally
                    {
                        //SQLCon.Close();
                    }
                }

            }

            foreach (Member m in lstBadList)
            {
                if (m.iID > 0)
                {
                    //need to check but it looks like is4c checks 2 things for Working Member Discoount.....SSI (hours worked) and isStaff==6
                    //aslong as hours is updated IS$C will do all the work. I'm keeping this BadList update seperate for now just in case the SQL statement needs to change
                    string sql = "UPDATE members SET Active=1,isStaff=6,SSI=" + m.iHours.ToString() + " WHERE CardNo= " + m.iID.ToString() + ";";

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
                                //updated more then one row
                            }
                            else if (rows < 1)
                            {
                                //updated more then one row
                            }


                        }
                    }
                    catch (Exception ex)
                    {

                        //there be errors 
                        errors++;
                    }
                    finally
                    {
                        //SQLCon.Close();
                    }

                }
            }
            if(errors==0)
                MessageBox.Show("You have updated the Member DB");
            else
                MessageBox.Show("You have updated the Member DB, but there were some errors");

        }


        //save the new connection string
        private void button3_Click(object sender, EventArgs e)
        {
            coop.Default.DBConnection = tbConnectionString.Text;
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
                    var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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

                    System.IO.File.Delete(@"temp.xlsx");

                    if (download_good)
                    {
                        FileStream file2 = new FileStream("temp.xlsx", FileMode.Create, FileAccess.Write);
                        stream2.WriteTo(file2);
                        file2.Close();
                        stream2.Close();

                        tbFilePath.Text = Application.StartupPath + @"\temp.xlsx";
                        ProcessFile();
                    }
                    else
                        MessageBox.Show("Well that didnt work. Download it yoursef.\nDownload/Export as an Excel Spreadsheet and drag it into this app.", "Error Downloading file", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch
                {
                    MessageBox.Show("Well that didnt work. Download it yoursef.", "Error Downloading file", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                finally
                {

                }
                return;
            }
        }
    }
}