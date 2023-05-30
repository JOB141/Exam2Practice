using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO;
using System.Diagnostics;


namespace Exam2Practice
{
    public partial class frmHoliday : Form
    {
        private PrintDocument printDocument;
        private PrintDialog printDialog;
        //Set up Database Connection
        OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\jacko\\source\\repos\\Exam2Practice\\tblTravelPractice.accdb");
        // Set up Database Adapter
        OleDbDataAdapter adapter;
        //Set up Data Table 
        DataTable dt = new DataTable();
        //Set up Command for Running SQL Queries
        OleDbCommand cmd;
        //Start Database at row 0
        int pos = 0;
        public frmHoliday()
        {
            InitializeComponent();
        }

        private void frmHoliday_Load(object sender, EventArgs e)
        {
            //Fill Form with Database Entries on load
            string sql = "SELECT * FROM TravelPractice";
            adapter = new OleDbDataAdapter(sql, conn);
            adapter.Fill(dt);
            showData(pos);

        }

        public void showData(int index)
        {
            //Set textboxs from Data Table
            tbHolidayNo.Text = dt.Rows[index]["HolidayNo"].ToString();
            tbDestination.Text = dt.Rows[index]["Destination"].ToString();
            //Convert Number to Currency
            tbCost.Text = "€" + Convert.ToDouble(dt.Rows[index]["Cost"]).ToString("#,##,0.00");
            dbDepartureDate.Text = dt.Rows[index]["DepartureDate"].ToString();
            tbNoOfDays.Text = dt.Rows[index]["NoOfDays"].ToString();
            
            //Check Available set checkbox to True
            if (dt.Rows[index]["Available"].ToString() == "True")
            {
                cbAvailable.Checked = true;
            }
            //Set Check box to False 
            else
            {
                cbAvailable.Checked = false;
            }

            //Set the Bottom textbox to total number of entries
            tbTotalEntries.Text = pos + 1 + " of " + dt.Rows.Count;
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            pos = 0;
            showData(pos);
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            pos--;
            if(pos >= 0)
            {
                showData(pos);
            }
            else
            {
                MessageBox.Show("END");
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            pos++;
            if (pos < dt.Rows.Count)
            {
                showData(pos);
            }
            else
            {
                MessageBox.Show("END");
                pos = dt.Rows.Count - 1;
            }
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            pos = dt.Rows.Count - 1;
            showData(pos);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (!ValidateHolidayNo())
            {
                return; // Validation failed, stop further execution
            }
            //Setup SQL Query
            string sql = "UPDATE TravelPractice SET [Destination] = ?, [Cost] = ?, [DepartureDate] = ?, [NoOfDays] = ?, [Available] = ? WHERE [HolidayNo] = ? ";
            //Start command
            cmd = new OleDbCommand(sql, conn);

            //Sets paramters in [] to the variable next to it IN ORDER,
            //e.g [Destination] = @Destination and tbDestination.Text = 2nd ?
            //Must have parameters in the same order you put them in Query
            cmd.Parameters.AddWithValue("@Destination", tbDestination.Text);
            cmd.Parameters.AddWithValue("@Cost", tbCost.Text);
            cmd.Parameters.AddWithValue("@DepartureDate", dbDepartureDate.Text);
            cmd.Parameters.AddWithValue("@NoOfDays", tbNoOfDays.Text.Remove(0, 1));
            cmd.Parameters.AddWithValue("@Available", cbAvailable.Checked);
            cmd.Parameters.AddWithValue("@HolidayNo", tbHolidayNo.Text);


            ExecuteCommand();
            //showData(pos);
        }

        private void ExecuteCommand()
        {
            try
            {
                //Open Connection to Database
                conn.Open();
                //Run the Query
                cmd.ExecuteNonQuery();
                //Close the Database connection
                conn.Close();
                //Update the Program 

                dt.Reset();
                adapter.Fill(dt);
                showData(pos);
            }
            catch (Exception ex)
            {
                conn.Close();
                MessageBox.Show($"Sorry there has been an ERROR{Environment.NewLine}{ex.Message}");
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateHolidayNo())
            {
                return; // Validation failed, stop further execution
            }
            //Set up SQL Query
            string sql = "INSERT INTO TravelPractice ([HolidayNo],[Destination],[Cost],[DepartureDate],[NoOfDays]) VALUES (?,?,?,?,?)";
            //Start Command
            cmd = new OleDbCommand(sql, conn);

            //Set parameters for Query, See Above for Explination
            cmd.Parameters.AddWithValue("@HolidayNo", tbHolidayNo.Text);
            cmd.Parameters.AddWithValue("@Destination", tbDestination.Text);
            cmd.Parameters.AddWithValue("@Cost", tbCost.Text.Trim('€'));
            cmd.Parameters.AddWithValue("@DepartureDate", dbDepartureDate.Text);
            cmd.Parameters.AddWithValue("@NoOfDays", tbNoOfDays.Text);
            cmd.Parameters.AddWithValue("@Available", cbAvailable.Checked);

            //Run the Query
            ExecuteCommand();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //Set up SQL Query
            string sql = "DELETE FROM TravelPractice WHERE [HolidayNo] = ?";
            //Start Command
            cmd = new OleDbCommand(sql, conn);

            //Set Parameter
            cmd.Parameters.AddWithValue("@HolidayNo", tbHolidayNo.Text);

            //Run Query
            ExecuteCommand();
           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
                ClearTextBoxes();
            }

            private void ClearTextBoxes()
            {
                tbDestination.Text = string.Empty;
                tbCost.Text = string.Empty;
                dbDepartureDate.Text = string.Empty;
                tbNoOfDays.Text = string.Empty;
                cbAvailable.Checked = false;
                tbHolidayNo.Text = string.Empty;
            }

        private bool ValidateHolidayNo()
        {
            int holidayNo;
            if (int.TryParse(tbHolidayNo.Text, out holidayNo))
            {
                if (holidayNo >= 200 && holidayNo <= 1000)
                {
                    return true; // HolidayNo is valid
                }
            }

            MessageBox.Show("Holiday number must be in the range 200 to 1000.", "Validation Error");
            return false; // HolidayNo is not valid
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            printDialog = new PrintDialog();
            printDocument = new PrintDocument();

            // Set up event handler for PrintPage
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                printDocument.Print();
            }
        }
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font fontHeader = new Font("Arial", 12, FontStyle.Bold);
            Font fontContent = new Font("Arial", 10);

            int startX = 50;
            int startY = 50;
            int offsetY = 20;

            // Print the header
            e.Graphics.DrawString("Downton Travel Page 01", fontHeader, Brushes.Black, startX, startY);
            e.Graphics.DrawString("Date " + DateTime.Now.ToString("dd/MM/yyyy"), fontHeader, Brushes.Black, startX, startY + offsetY);

            // Print the column headers
            string columnHeader = "Holiday No   Destination                 Departure Date   Cost       Available";
            e.Graphics.DrawString(columnHeader, fontHeader, Brushes.Black, startX, startY + offsetY * 3);

            // Create a StringBuilder to store the content for the .txt file
            StringBuilder sb = new StringBuilder();

            // Append the header to the StringBuilder
            sb.AppendLine("Downton Travel Page 01");
            sb.AppendLine("Date " + DateTime.Now.ToString("dd/MM/yyyy"));
            sb.AppendLine(columnHeader);

            // Print the column headers
            e.Graphics.DrawString("Holiday No   Destination                 Departure Date   Cost       Available", fontHeader, Brushes.Black, startX, startY + offsetY * 3);

            // Retrieve and print the records from the database
            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\jacko\\source\\repos\\Exam2Practice\\tblTravelPractice.accdb"))
            {
                conn.Open();
                string sql = "SELECT HolidayNo, Destination, DepartureDate, Cost, Available FROM TravelPractice";
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                int currentRecord = 1;

                while (reader.Read())
                {
                    string holidayNo = reader["HolidayNo"].ToString();
                    string destination = reader["Destination"].ToString();
                    string departureDate = reader["DepartureDate"].ToString();
                    string cost = reader["Cost"].ToString();
                    string available = reader.GetBoolean(reader.GetOrdinal("Available")) ? "Yes" : "No";

                    // Format the record and print
                    string record = $"{holidayNo,-11} {destination,-27} {departureDate,-15} {cost,-10} {available}";
                    e.Graphics.DrawString(record, fontContent, Brushes.Black, startX, startY + offsetY * (currentRecord + 4));

                    // Append the record to the StringBuilder
                    sb.AppendLine(record);

                    currentRecord++;
                }

                reader.Close();
                conn.Close();
            }

            // Save the content to a .txt file
            string filePath = @"C:\Users\jacko\source\repos\Exam2Practice\output.txt";
            File.WriteAllText(filePath, sb.ToString());

            // Open the .txt file using the default associated program
            System.Diagnostics.Process.Start(filePath);
        }


    }
}

