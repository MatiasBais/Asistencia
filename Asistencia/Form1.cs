using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Configuration;
using Asistencia.Properties;
using MySql.Data.MySqlClient;
namespace Asistencia
{
    public partial class Form1 : Form
    {       
        public Form1()
        {
            InitializeComponent();
        }
        private string cadenaconexion() 
        {
            string caca = "data source=localhost;database=asistencia3;user id = root; password=root";
           return caca;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(cadenaconexion());
            string consulta = "select * from Alumno where curso='" + comboBox1.Text + "' order by nombre";
                MySqlDataAdapter data = new MySqlDataAdapter(consulta, conn);
                DataTable ds = new DataTable();
                data.Fill(ds);
                dataGridView1.DataSource = ds;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[4].Visible = false;     
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select * from alumno where curso='" + comboBox1.Text + "' order by nombre";
            MySqlDataAdapter data = new MySqlDataAdapter(consulta, cnn);
            DataTable ds = new DataTable();
            data.Fill(ds);
            dataGridView1.DataSource = ds;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            numericUpDown1.Value = id;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            using (MySqlConnection conn = new MySqlConnection(cadenaconexion()))
            {
                string query = "SELECT * FROM alumno WHERE idalumno =" + numericUpDown1.Value;
                MySqlCommand cmd = new MySqlCommand(query, conn);
                conn.Open();
                MySqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    t_nombre.Text = Convert.ToString(reader["Nombre"]);
                    t_tel.Text = Convert.ToString(reader["Telefono"]);
                    t_dni.Text = Convert.ToString(reader["DNI"]);
                    comboBox1.Text = Convert.ToString(reader["curso"]);
                }
                else
                {
                    t_nombre.Text = "";
                    t_dni.Text = "";
                    t_tel.Text = "";
                }
                button1.Visible=false;
                button2.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string nombre, dni, telefono, curso;
            nombre = t_nombre.Text;
            dni = t_dni.Text;
            telefono = t_tel.Text;
            curso = comboBox1.Text;
            string consulta = "INSERT INTO alumno (Nombre, DNI, Telefono, Curso) VALUES ('" + nombre + "','" + dni + "','" + telefono + "','" + curso + "')";
            cnn.Open();
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
            string curso2 = comboBox1.Text;
            string consulta2 = "select * from alumno where Curso ='" + curso2 + "' order by nombre";
            MySqlDataAdapter data2 = new MySqlDataAdapter(consulta2, cnn);
            DataTable dtt2 = new DataTable();
            data2.Fill(dtt2);
            dataGridView1.DataSource = dtt2;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            t_nombre.Clear();
            t_dni.Clear();
            t_tel.Clear();
            //MessageBox.Show("hola");
            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string nombre, dni, telefono, curso;
                nombre = t_nombre.Text;
                dni = t_dni.Text;
                telefono = t_tel.Text;
                curso = comboBox1.Text;
                string consulta = "UPDATE alumno SET Nombre='" + nombre + "',DNI='" + dni + "',Telefono='" + telefono + "',Curso='" + curso + "' WHERE idAlumno =" + numericUpDown1.Value;
                cnn.Open();
                MySqlCommand cmd = new MySqlCommand(consulta, cnn);
                cmd.ExecuteNonQuery();
                cnn.Close();
                string curso2 = comboBox1.Text;
                string consulta2 = "select * from alumno where Curso ='" + curso2 + "' order by nombre";
                MySqlDataAdapter data2 = new MySqlDataAdapter(consulta2, cnn);
                DataTable dtt2 = new DataTable();
                data2.Fill(dtt2);
                dataGridView1.DataSource = dtt2;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[4].Visible = false;
                numericUpDown1.Value = 0;
                t_nombre.Clear();
                t_dni.Clear();
                t_tel.Clear();
                button1.Visible = true;
                button2.Visible = false;
                button3.Visible = false;
            button4.Visible = false;
            }
        private void button3_Click(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "DELETE FROM alumno WHERE idalumno=" + numericUpDown1.Value;
                cnn.Open();
                MySqlCommand cmd = new MySqlCommand(consulta, cnn);
                cmd.ExecuteNonQuery();
                cnn.Close();
                string curso2 = comboBox1.Text;
                string consulta2 = "select * from alumno where Curso ='" + curso2 + "' order by nombre";
                MySqlDataAdapter data2 = new MySqlDataAdapter(consulta2, cnn);
                DataTable dtt2 = new DataTable();
                data2.Fill(dtt2);
                dataGridView1.DataSource = dtt2;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[4].Visible = false;
                numericUpDown1.Value = 0;
                t_nombre.Clear();
                t_dni.Clear();
                t_tel.Clear();
                button1.Visible = true;
                button2.Visible = false;
                button3.Visible = false;
            button4.Visible = false;
        }

        private void t_tel_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            t_nombre.Clear();
            t_dni.Clear();
            t_tel.Clear(); 
            button1.Visible = true;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            
        }
        }
    }

