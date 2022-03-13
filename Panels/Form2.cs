using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace Panels
{
    public partial class frmRegistration : Form
    {
        //string connection
        string path = @"Data Source=LAPTOP-67FGJG2C\MSSQLSERVER1;Initial Catalog=registration;Persist Security Info=True;User ID=testForm;Password=testForm";

        SqlConnection con = new SqlConnection();
        SqlDataAdapter adpt = new SqlDataAdapter();
        System.Data.DataTable dt = new System.Data.DataTable();
        int ID;

        public frmRegistration()
        {
            InitializeComponent();
            con = new SqlConnection(path);
            display();
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtName.Text == "" || txtSurname.Text == "" || txtDesignation.Text == "" || txtEmail.Text == "" || txtID.Text == "" || txtAddress.Text == "")
            {
                MessageBox.Show("Please Fill the Empty fields");
            }
            else
            {
                string gender = "Male";
                if (rbtnFemale.Checked == true)
                    gender = "Female";


                try
                {


                    con.Open();

                    SqlCommand cmd = new SqlCommand("insert into Employee (Employee_Name,Employee_Surname,Employee_Designation,Employee_Email,Emp_ID,Gender,Address) values ('" + txtName.Text + "', '" + txtSurname.Text + "', '" + txtDesignation.Text + "','" + txtEmail.Text + "','" + txtID.Text + "', '" + gender + "', '" + txtAddress.Text + "')", con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Your Data have been saved into the database");
                    clean();
                    display();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void clean()
        {
            txtName.Text = "";
            txtSurname.Text = "";
            txtDesignation.Text = "";
            txtEmail.Text = "";
            txtID.Text = "";
            txtAddress.Text = "";
            rbtnMale.Checked = false;
            rbtnFemale.Checked = false;
        }
        public void display()
        {
            try
            {
                dt = new System.Data.DataTable();
                con.Open();
                adpt = new SqlDataAdapter("select * from Employee", con);
                adpt.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();

            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    


        
        
        
        
        

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try 
            {
                string gender = "Male";
                if (rbtnFemale.Checked == true)
                    gender = "Female";

                con.Open();
                SqlCommand cmd = new SqlCommand("update employee set Employee_Name='" + txtName.Text + "', Employee_Surname='" + txtSurname.Text + "', Employee_Designation='" + txtDesignation.Text + "',Employee_Email='" + txtEmail.Text + "', Emp_ID='" + txtID.Text + "', Gender='" + gender + "', Address='" + txtAddress.Text + "' where Employee_Id='" + ID + "' ", con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Your data have been updated!");
                display();
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("delete from Employee where Employee_Id='" + ID + "'",con);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Your record has been deleted");
                display();

            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ID=int.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtSurname.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtDesignation.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtEmail.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            txtID.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();

            rbtnMale.Checked = true;
            rbtnFemale.Checked = false;

            if (dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString()=="Female")
            {
                rbtnMale.Checked=false;
                rbtnFemale.Checked = true;
            }
            
            txtAddress.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            btnUpdate.Enabled=true;
            btnDelete.Enabled=true;    
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            con.Open();
            adpt = new SqlDataAdapter("select * from employee where Employee_Name like '%" + txtSearch.Text + "%' ", con);
            dt = new System.Data.DataTable();
            adpt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                Microsoft.Office.Interop.Excel.Application Excell = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = Excell.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)Excell.ActiveSheet;
                Excell.Visible = true;

                for (int j = 2; j <= dataGridView1.Rows.Count; j++)
                {
                    for (int i = 1; i <= 1; i++)
                    {
                        ws.Cells[j, i] = dataGridView1.Rows[j - 2].Cells[i - 1].Value;
                    }
                }

                for (int i = 1; i < dataGridView1.Columns.Count+1; i++)
                {
                    ws.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Columns.Count-1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        ws.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }




            } catch (Exception) 
            { 
            }



            
        }

        private void frmRegistration_Load(object sender, EventArgs e)
        {

        }
    }
}
