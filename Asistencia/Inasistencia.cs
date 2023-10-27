using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Configuration;
using MySql.Data.MySqlClient;
using Asistencia.Properties;
namespace Asistencia
{
    public partial class Inasistencia : Form
    {
        public Inasistencia()
        {
            InitializeComponent();
        }
        private string cadenaconexion()
        {
            string caca = "database=asistencia3;data source=localhost; user id=root";
            return caca;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 alumnos = new Form1();
            alumnos.Show();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Inasistencias ina = new Inasistencias();
            ina.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ina.Show();
        }
        private void button3_Click(object sender, EventArgs e)
        {
           
        }
        private void inasistenciasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Inasistencias ina = new Inasistencias();
            ina.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ina.Show();
        }
        private void alumnosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void porNombreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AlumnoNombre al = new AlumnoNombre();
            al.Show();
        }
        private void porDniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AlumnoDNI al = new AlumnoDNI();
            al.Show();
        }
        private void pasarTodosDeGradoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            if (MessageBox.Show("Estás seguro?No hay vuelta atras", "My Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string consulta = "UPDATE alumno SET curso='Antiguo alumno' where curso='7mo info';UPDATE alumno SET curso='Antiguo alumno' where curso='7mo electro';UPDATE alumno SET curso='7mo info' where curso='6to info';UPDATE alumno SET curso='7mo electro' where curso='6to electro';UPDATE alumno SET curso='6to electro' where curso='5to electro';UPDATE alumno SET curso='6to info' where curso='5to info';UPDATE alumno SET curso='5to info' where curso='4to info';UPDATE alumno SET curso='5to electro' where curso='4to electro';UPDATE alumno SET curso='4to electro' where curso='3ro 1era';UPDATE alumno SET curso='4to info' where curso='3ro 2da';UPDATE alumno SET curso='3ro 1era' where curso='2do 1era';UPDATE alumno SET curso='3ro 2da' where curso='2do 2da';UPDATE alumno SET curso='2do 1era' where curso='1ero 1era';UPDATE alumno SET curso='2do 2da' where curso='1ero 2da';"; 
                cnn.Open();
                MySqlCommand cmd = new MySqlCommand(consulta, cnn);
                cmd.ExecuteNonQuery();
                cnn.Close();
                MessageBox.Show("Se ha subido de año a todos los alumnos, para los repitentes edite al curso correspondiente, lo mismo para el que se haya cambiado de division. Con respecto al pase de 3er año a 4to, por defecto 3ero 1era pasa a electro y 2da a info, haga los cambios correspondientes en Alumnos.");
                MessageBox.Show("No obstante el alumnado que finalizo sus estudios pasa a la categoria Antiguo Alumno la cual tambien puede ser editado en caso de que haya algun repitente");
            }
            
        }
        int dias = 10;
        private void Inasistencia_Load(object sender, EventArgs e)
        {
            
           
                MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                string query = "select * from alumno";
                MySqlCommand cmd = new MySqlCommand(query, cnn);
                MySqlDataAdapter adp = new MySqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adp.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Rows[0].Visible = true;
                
                for (int i = 1; i < dataGridView1.RowCount; i++)
                {
                    
                        cnn.Open();
                        int id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                        string query2 = "select sum(valor) as total from inasistencia where idalumno=" + id;
                        MySqlCommand cmd2 = new MySqlCommand(query2, cnn);
                        MySqlDataReader reader = cmd2.ExecuteReader();
                        if (reader.Read())
                        {
                            bool asd = Convert.IsDBNull(reader["total"]);
                            if (asd == true)
                            {
                                dataGridView1.Rows[i].Visible = false;
                            }
                            else if (Convert.ToDouble(reader["total"]) < dias)
                            {
                                dataGridView1.Rows[i].Visible = false;
                            }


                        }
                        reader.Close();
                        cnn.Close();
                    
                    
                }
                try
                {
                    dataGridView1.Rows[0].Visible = false;
                }
                catch
                {
                }
            }
            
        

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
                dias = 10;

            dataGridView1.Rows[0].Visible = true;

            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            for (int i = 1; i < dataGridView1.RowCount; i++)
            {

                cnn.Open();
                int id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                string query2 = "select sum(valor) as total from inasistencia where idalumno=" + id;
                MySqlCommand cmd2 = new MySqlCommand(query2, cnn);
                MySqlDataReader reader = cmd2.ExecuteReader();
                if (reader.Read())
                {
                    bool asd = Convert.IsDBNull(reader["total"]);
                    if (asd == true)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else if (Convert.ToDouble(reader["total"]) < dias)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }


                }
                reader.Close();
                cnn.Close();


            }
            try
            {
                dataGridView1.Rows[0].Visible = false;
            }
            catch
            {
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
                dias = 15;

            dataGridView1.Rows[0].Visible = true;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());

            for (int i = 1; i < dataGridView1.RowCount; i++)
            {

                cnn.Open();
                int id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                string query2 = "select sum(valor) as total from inasistencia where idalumno=" + id;
                MySqlCommand cmd2 = new MySqlCommand(query2, cnn);
                MySqlDataReader reader = cmd2.ExecuteReader();
                if (reader.Read())
                {
                    bool asd = Convert.IsDBNull(reader["total"]);
                    if (asd == true)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else if (Convert.ToDouble(reader["total"]) < dias)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else 
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }


                }
                reader.Close();
                cnn.Close();


            }
            try
            {
                dataGridView1.Rows[0].Visible = false;
            }
            catch
            {
            }
        }


        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
                dias = 20;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());

            dataGridView1.Rows[0].Visible = true;
            for (int i = 1; i < dataGridView1.RowCount; i++)
            {

                cnn.Open();
                int id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                string query2 = "select sum(valor) as total from inasistencia where idalumno=" + id;
                MySqlCommand cmd2 = new MySqlCommand(query2, cnn);
                MySqlDataReader reader = cmd2.ExecuteReader();
                if (reader.Read())
                {
                    bool asd = Convert.IsDBNull(reader["total"]);
                    if (asd == true)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else if (Convert.ToDouble(reader["total"]) < dias)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }

                }
                reader.Close();
                cnn.Close();


            }
            try
            {
                dataGridView1.Rows[0].Visible = false;
            }
            catch
            {
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                dias = 28;
            }


            dataGridView1.Rows[0].Visible = true;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            for (int i = 1; i < dataGridView1.RowCount; i++)
            {

                cnn.Open();
                int id = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                string query2 = "select sum(valor) as total from inasistencia where idalumno=" + id;
                MySqlCommand cmd2 = new MySqlCommand(query2, cnn);
                MySqlDataReader reader = cmd2.ExecuteReader();
                if (reader.Read())
                {
                    bool asd = Convert.IsDBNull(reader["total"]);
                    if (asd == true)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else if (Convert.ToDouble(reader["total"]) < dias)
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                }
                reader.Close();
                cnn.Close();
            }
            try
            {
                dataGridView1.Rows[0].Visible = false;
            }
            catch
            {
            }
        }
    }
}
