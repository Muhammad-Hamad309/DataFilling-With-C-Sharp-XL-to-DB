using FillDataFromXL.Logic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;

namespace FillDataFromXL
{
    public partial class Form1 : Form
    {
        string DataPath = "YourDataPath.xlsx";

        FileReader file = new FileReader();

        SqlConnection con = new SqlConnection("Data Source= ;Initial Catalog= ;Integrated Security=True");

        public bool UploadToDb(string name, string code)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("Insert into Port(Name,Code) Values('" + name + "','" + code + "')", con);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int Failed = 0;
            // Country Data
            var CountryDt = file.ReadExcelFile(DataPath, 0 /*This is the Sheet Number*/);
            var lstDataRow = CountryDt.Rows.Cast<DataRow>().ToList();
            foreach (var item in lstDataRow)
            {
                string name = item["Name"].ToString();
                string IsoCode = item["Code"].ToString();
                bool stat = UploadToDb(name, IsoCode);

                if (stat == false) { Failed++; }
            }

            //SQL Bulk Upload
            var CountryDt1 = file.ReadExcelFile(DataPath, 0/*This is the Sheet Number*/);
            CountryDt1.Columns.Add("Name");
            CountryDt1.Columns.Add("Code");

            //Add Columns for Default Data or Common Data

            foreach (DataRow row in CountryDt1.Rows)
            {
                //Fill Columns with Common or Default Value
            }

            CountryDt1.AcceptChanges();

            SqlBulkCopy sbc = new SqlBulkCopy(con);

            sbc.DestinationTableName = "PaymentTerm";
            sbc.ColumnMappings.Add("Name" /*This is the Column Name of the XL File should Match as in File*/, "Name");
            sbc.ColumnMappings.Add("Code" /*This is the Column Name of the XL File should Match as in File*/, "Code");

            con.Open();
            sbc.WriteToServer(CountryDt1);
            con.Close();
            
            dataGridView1.DataSource = CountryDt;
            //dataGridView2.DataSource = CountryDt1;
            MessageBox.Show("Total Failed " + Failed);

        }
    }
}