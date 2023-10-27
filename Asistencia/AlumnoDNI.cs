using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Asistencia.Properties;
using System.Configuration;
using System.Data.Sql;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace Asistencia
{
    public partial class AlumnoDNI : Form
    {
        public AlumnoDNI()
        {
            InitializeComponent();
        }

        private string cadenaconexion()
        {
            string caca = "database=asistencia3;data source=localhost; user id=root";
            return caca;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select Nombre, Curso, Telefono, DNI from alumno where DNI LIKE '%" + textBox1.Text + "%' and curso !='Antiguo Alumno' and curso !=''";
            MySqlDataAdapter data = new MySqlDataAdapter(consulta, cnn);
            DataTable dtt = new DataTable();
            data.Fill(dtt);
            dataGridView1.DataSource = dtt;
        }
        private void AlumnoDNI_Load(object sender, EventArgs e)
        {
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select Nombre, Curso, Telefono, DNI from alumno where curso !='Antiguo Alumno' and curso !=''";
            MySqlDataAdapter data = new MySqlDataAdapter(consulta, cnn);
            DataTable dtt = new DataTable();
            data.Fill(dtt);
            dataGridView1.DataSource = dtt;
        }
    }
}
