using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Configuration;
using MySql.Data.MySqlClient;
using Asistencia.Properties;

namespace Asistencia
{
    public partial class Inasistencias : Form
    {
        public Inasistencias()
        {
            InitializeComponent();
        }
        private string cadenaconexion()
        {
            string caca = "database=asistencia3;data source=localhost; user id=root";
            return caca;
        }
        private string Mes()
        {
            string between;
            string mess = c_mes.Text;
            switch (mess)
            {
                case "Marzo":
                    between = "'" + comboBox12.Text + "/02/28' and '" + comboBox12.Text + "/03/31'";
                    break;
                case "Abril":
                    between = "'" + comboBox12.Text + "/03/31' and '" + comboBox12.Text + "/04/30'";
                    break;
                case "Mayo":
                    between = "'" + comboBox12.Text + "/04/30' and '" + comboBox12.Text + "/05/31'";
                    break;
                case "Junio":
                    between = "'" + comboBox12.Text + "/05/31' and '" + comboBox12.Text + "/06/30'";
                    break;
                case "Julio":
                    between = "'" + comboBox12.Text + "/06/30' and '" + comboBox12.Text + "/07/31'";
                    break;
                case "Agosto":
                    between = "'" + comboBox12.Text + "/07/31' and '" + comboBox12.Text + "/08/31'";
                    break;
                case "Septiembre":
                    between = "'" + comboBox12.Text + "/09/01' and '" + comboBox12.Text + "/09/30'";
                    break;
                case "Octubre":
                    between = "'" + comboBox12.Text + "/09/30' and '" + comboBox12.Text + "/10/31'";
                    break;
                case "Noviembre":
                    between = "'" + comboBox12.Text + "/10/31' and '" + comboBox12.Text + "/11/30'";
                    break;
                default:
                    int añosig = Convert.ToInt32(comboBox12.Text) + 1;
                    between = "'" + comboBox12.Text + "/11/30' and '" + añosig + "/01/01'";
                    break;
            }
            return between;
        }
        private int dia()
        {
            string mess = c_mes.Text;
            int dias;
            switch (mess)
            {
                case "Marzo":
                    dias = 31;
                    break;
                case "Abril":
                    dias = 30;
                    break;
                case "Mayo":
                    dias = 31;
                    break;
                case "Junio":
                    dias = 30;
                    break;
                case "Julio":
                    dias = 31;
                    break;
                case "Agosto":
                    dias = 31;
                    break;
                case "Septiembre":
                    dias = 30;
                    break;
                case "Octubre":
                    dias = 31;
                    break;
                case "Noviembre":
                    dias = 30;
                    break;
                default:
                    dias = 31;
                    break;
            }
            return dias;
        }
        private string nombredia(int i)
        {
            int nmes = c_mes.SelectedIndex + 3;
            DateTime dateValue = new DateTime(Convert.ToInt32(comboBox12.Text), nmes, i);
            string diames;
            switch (Convert.ToString(dateValue.DayOfWeek))
            {
                case "Monday":
                    diames = "L";
                    break;
                case "Tuesday":
                    diames = "M";
                    break;
                case "Wednesday":
                    diames = "M";
                    break;
                case "Thursday":
                    diames = "J";
                    break;
                case "Friday":
                    diames = "V";
                    break;
                case "Saturday":
                    diames = "S";
                    break;
                default:
                    diames = "D";
                    break;

            }
            return diames;
        }
        private void CargarDatos()
        {

            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select * from Alumno where curso ='" + comboBox3.Text + "' order by Nombre";
            cnn.Open();
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            MySqlDataReader reader = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Nombre");
            string nombrecolumna;
            for (int i = 1; i <= dia(); i++)
            {
                int nmes = c_mes.SelectedIndex + 3;
                nombrecolumna = i.ToString() + nombredia(i);
                DateTime dateValue = new DateTime(Convert.ToInt32(comboBox12.Text), nmes, i);
                dt.Columns.Add(nombrecolumna);
                dt.Columns[nombrecolumna].AllowDBNull = false;
                dt.Columns[nombrecolumna].DefaultValue = "P";
            }
            while (reader.Read())
            {
                DataRow row = dt.NewRow();
                row["Id"] = reader["idAlumno"].ToString();
                row["Nombre"] = reader["Nombre"].ToString();
                dt.Rows.Add(row);
            }
            reader.Close();
            cnn.Close();
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].MinimumWidth = 150;

            for (int i = 1; i <= dia(); i++)
            {
                int nmes = c_mes.SelectedIndex + 3;
                nombrecolumna = i.ToString() + nombredia(i);
                DateTime dateValue = new DateTime(Convert.ToInt32(comboBox12.Text), nmes, i);
                if (Convert.ToString(dateValue.DayOfWeek) == "Saturday" || Convert.ToString(dateValue.DayOfWeek) == "Sunday")
                {
                    dataGridView1.Columns[nombrecolumna].Visible = false;
                }
            }
            int nmes2 = c_mes.SelectedIndex + 3;

            MySqlConnection conn = new MySqlConnection(cadenaconexion());
            string query = "SELECT * FROM inasistencia, alumno WHERE inasistencia.idalumno = alumno.idalumno and fecha between " + Mes() + " and Curso ='" + comboBox3.Text + "' and inasistencia.Turno='" + comboBox4.Text + "'";
            MySqlCommand cmd2 = new MySqlCommand(query, conn);
            conn.Open();
            MySqlDataReader reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            {
                double valor = Convert.ToDouble(reader2["Valor"]);
                int idalumno = Convert.ToInt32(reader2["idAlumno"]);
                string fech = reader2["Fecha"].ToString();
                DateTime sas = new DateTime();
                sas = Convert.ToDateTime(fech);
                int diia = sas.Day + 1;
                string justif = Convert.ToString(reader2["Justif"]);
                int roww = dataGridView1.RowCount;
                for (int y = 0; y < roww; y++)
                {
                    if (idalumno == Convert.ToInt32(dataGridView1.Rows[y].Cells[0].Value))
                    {
                        dataGridView1.Rows[y].Cells[diia].Value = valor;

                        if (justif == "Si")
                        {
                            if (comboBox2.Text == "Verde")
                            {
                                dataGridView1.Rows[y].Cells[diia].Style.BackColor = Color.Green;
                            }
                            else
                            {
                                dataGridView1.Rows[y].Cells[diia].Style.BackColor = Color.LightBlue;
                            }
                        }
                        else
                        {
                            if (comboBox1.Text == "Rosa")
                            {
                                dataGridView1.Rows[y].Cells[diia].Style.BackColor = Color.Pink;
                            }
                            else
                            {
                                dataGridView1.Rows[y].Cells[diia].Style.BackColor = Color.Red;
                            }
                        }
                    }

                }
            }






        }
        private void Inasistencias_Load(object sender, EventArgs e)
        {
            c_mes.SelectedIndex = Convert.ToInt32(DateTime.Today.Month) - 3;
            numericUpDown1.Value = Convert.ToInt32(DateTime.Today.Year);
            CargarDatos();

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarDatos();
            textBox9.Text = "";
        }
        private void c_mes_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarDatos();
            textBox9.Text = "";
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarDatos();
            textBox9.Text = "";
        }
        private int idfalta(int idalumno, string fecha, string turno)
        {

            MySqlConnection conn = new MySqlConnection(cadenaconexion());

            string query = "SELECT * FROM inasistencia where idalumno =" + idalumno + " and fecha='" + fecha + "' and Turno='" + turno + "'";
            MySqlCommand cmd = new MySqlCommand(query, conn);
            conn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            int idfalta;
            if (reader.Read())
            {
                idfalta = Convert.ToInt32(reader["idfalta"]);
            }
            else
            {
                idfalta = 0;
            }
            return idfalta;
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Convert.ToString(dataGridView1.CurrentCell.Value) == "P" || Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
            {
                try
                {
                    int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                    int nmes = c_mes.SelectedIndex + 3;
                    int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                    string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                    string query;
                    string turno = comboBox4.Text;

                    if (Convert.ToString(dataGridView1.CurrentCell.Value) == "P")
                    {
                        query = "insert into inasistencia (idalumno, fecha, valor, Justif, Turno) values (" + idalumno + ", '" + fecha + "', 1, 'No','" + comboBox4.Text + "')";
                        dataGridView1.CurrentCell.Value = "1";
                        if (comboBox1.Text == "Rojo")
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                        }
                    }
                    else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1")
                    {
                        query = "Update inasistencia set valor = 0.75 where idfalta =" + idfalta(idalumno, fecha, turno);
                        dataGridView1.CurrentCell.Value = "0,75";
                    }
                    else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75")
                    {
                        query = "Update inasistencia set valor = 0.5 where idfalta =" + idfalta(idalumno, fecha, turno);
                        dataGridView1.CurrentCell.Value = "0,5";
                    }
                    else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5")
                    {
                        query = "Update inasistencia set valor = 0.25 where idfalta =" + idfalta(idalumno, fecha, turno);
                        dataGridView1.CurrentCell.Value = "0,25";
                    }
                    else
                    {
                        query = "delete from inasistencia where idfalta=" + idfalta(idalumno, fecha, turno);
                        dataGridView1.CurrentCell.Value = "P";
                        dataGridView1.CurrentCell.Style.BackColor = Color.White;

                        textBox9.Enabled = false;
                    }
                    MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                    MySqlCommand cmd = new MySqlCommand(query, cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    totalfaltas();
                    faltasinjus();
                    faltasjusti();
                    faltasaño();
                    justiaño();
                    injustiaño();
                    observaciones();
                    // CargarDatos();
                }
                catch (Exception el)
                {
                    MessageBox.Show(el + "");
                }
            }
        }
        private void totalfaltas()
        {
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + Mes();
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox3.Text = reader["total"].ToString();
                if (textBox3.Text == "")
                {
                    textBox3.Text = "0";
                }
            }

            cnn.Close();
        }
        private void faltasinjus()
        {
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + Mes() + " and Justif != 'Si'";
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox2.Text = reader["total"].ToString();
                if (textBox2.Text == "")
                {
                    textBox2.Text = "0";
                }
            }

            cnn.Close();
        }
        private void faltasjusti()
        {
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + Mes() + " and Justif='Si'";
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox1.Text = reader["total"].ToString();
                if (textBox1.Text == "")
                {
                    textBox1.Text = "0";
                }

            }

            cnn.Close();
        }
        private void faltasaño()
        {
            int año1 = Convert.ToInt32(comboBox12.Text) - 1;
            int año2 = Convert.ToInt32(comboBox12.Text) + 1;
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = "'" + año1 + "/12/31' and '" + año2 + "/01/01'";
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + fecha;
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox4.Text = reader["total"].ToString();
                if (textBox4.Text == "")
                {
                    textBox4.Text = "0";
                }
            }

            cnn.Close();
        }
        private void justiaño()
        {
            int año1 = Convert.ToInt32(comboBox12.Text) - 1;
            int año2 = Convert.ToInt32(comboBox12.Text) + 1;
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = "'" + año1 + "/12/31' and '" + año2 + "/01/01'";
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + fecha + " and Justif='Si'";
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox6.Text = reader["total"].ToString();
                if (textBox6.Text == "")
                {
                    textBox6.Text = "0";
                }
            }

            cnn.Close();
        }
        private void injustiaño()
        {
            int año1 = Convert.ToInt32(comboBox12.Text) - 1;
            int año2 = Convert.ToInt32(comboBox12.Text) + 1;
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = "'" + año1 + "/12/31' and '" + año2 + "/01/01'";
            MySqlConnection cnn = new MySqlConnection(cadenaconexion());
            string consulta = "select sum(valor) as total from inasistencia where idalumno=" + idalumno + " and fecha between " + fecha + " and Justif!= 'Si'";
            MySqlCommand cmd = new MySqlCommand(consulta, cnn);
            cnn.Open();
            MySqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                textBox5.Text = reader["total"].ToString();
                if (textBox5.Text == "")
                {
                    textBox5.Text = "0";
                }
            }

            cnn.Close();
        }
        private void observaciones()
        {
            if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
            {
                int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                int nmes = c_mes.SelectedIndex + 3;
                int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                string turno = comboBox4.Text;
                MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                string consulta = "select observaciones from inasistencia where idfalta=" + idfalta(idalumno, fecha, turno); ;
                MySqlCommand cmd = new MySqlCommand(consulta, cnn);
                cnn.Open();
                MySqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    textBox9.Text = reader["observaciones"].ToString();
                }
                cnn.Close();
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
            {
                textBox9.ReadOnly = false;
                button1.Visible = true;
            }
            else
            {
                textBox9.ReadOnly = true;
                textBox9.Text = "";
                button1.Visible = false;
            }
            totalfaltas();
            faltasinjus();
            faltasjusti();
            faltasaño();
            justiaño();
            injustiaño();
            observaciones();
        }
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
            string query = "";
            string turno = comboBox4.Text;
            if (Convert.ToString(dataGridView1.CurrentCell.Value) != "P")
            {
                if (dataGridView1.CurrentCell.Style.BackColor == Color.LightBlue)
                {
                    query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                    if (comboBox1.SelectedIndex == 1)
                        dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                    else
                        dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                }
                else if (dataGridView1.CurrentCell.Style.BackColor == Color.Pink)
                {
                    query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                    if (comboBox2.SelectedIndex == 0)
                        dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                    else
                        dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                }
                else if (dataGridView1.CurrentCell.Style.BackColor == Color.Red)
                {
                    query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                    if (comboBox2.SelectedIndex == 0)
                        dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                    else
                        dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                }
                else if (dataGridView1.CurrentCell.Style.BackColor == Color.Green)
                {
                    query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                    if (comboBox1.SelectedIndex == 1)
                        dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                    else
                        dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                }
                try
                {
                    MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                    MySqlCommand cmd = new MySqlCommand(query, cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    totalfaltas();
                    faltasinjus();
                    faltasjusti();
                    faltasaño();
                    justiaño();
                    injustiaño();
                    observaciones();
                    //CargarDatos();
                }
                catch (Exception el)
                {
                    MessageBox.Show(el + "");
                }
            }
        }
        private void ExportarDataGridViewExcel(DataGridView dataGridView1)
        {
            SaveFileDialog fichero = new SaveFileDialog();
            fichero.Filter = "Excel (*.xls)|*.xls";
            if (fichero.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application aplicacion;
                Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libros_trabajo = aplicacion.Workbooks.Add();
                hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                string nombrecolumna;
                int j3 = 0;
                for (int i = 1; i < dataGridView1.Rows.Count; i++)
                {
                    int j2 = 0;

                    for (int j = 1; j <= dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i - 1].Cells[j - 1].Visible == true)
                        {
                            hoja_trabajo.Cells[i + 1, (j - j2)] = dataGridView1.Rows[i].Cells[j - 1].Value.ToString();
                        }
                        else
                        {
                            j2++;
                        }
                    }
                    if (dataGridView1.Rows[i - 1].Cells[i + 1].Visible == true)
                    {
                        int nmes;
                        nmes = c_mes.SelectedIndex + 3;
                        nombrecolumna = i.ToString() + nombredia(i);
                        DateTime dateValue = new DateTime(Convert.ToInt32(comboBox12.Text), nmes, i);
                        hoja_trabajo.Cells[1, i - j3] = nombrecolumna;
                    }
                    else
                    {
                        j3++;
                    }
                }
                hoja_trabajo.Cells[1, 1] = "Nombre";
                libros_trabajo.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                libros_trabajo.Close(true);
                OpenMicrosoftExcel(fichero.FileName);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExportarDataGridViewExcel(dataGridView1);
        }
        static void OpenMicrosoftExcel(string f)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "EXCEL.EXE";
            startInfo.Arguments = f;
            Process.Start(startInfo);
        }
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            int nmes = c_mes.SelectedIndex + 3;
            int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
            string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
            string query;
            string turno = comboBox4.Text;
            if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
            {
                query = "Update inasistencia set observaciones='" + textBox9.Text + "' where idfalta =" + idfalta(idalumno, fecha, turno);
                MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                MySqlCommand cmd = new MySqlCommand(query, cnn);
                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
            }

        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            comboBox12.Text = numericUpDown1.Value.ToString();
        }
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (dataGridView1.CurrentCell.ColumnIndex != 1)
                {
                    if (Convert.ToString(dataGridView1.CurrentCell.Value) != "P")
                    {
                        var relativeMousePosition = dataGridView1.PointToClient(Cursor.Position);
                        contextMenuStrip1.Show(dataGridView1, relativeMousePosition);
                    }
                }
            }
        }
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int row = dataGridView1.CurrentCell.RowIndex;
            int column = dataGridView1.CurrentCell.ColumnIndex;
            CargarDatos();
            dataGridView1.CurrentCell = dataGridView1.Rows[row].Cells[column];

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int row = dataGridView1.CurrentCell.RowIndex;
            int column = dataGridView1.CurrentCell.ColumnIndex;
            CargarDatos();
            dataGridView1.CurrentCell = dataGridView1.Rows[row].Cells[column];
        }
        private void justificarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
            {
                int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                int nmes = c_mes.SelectedIndex + 3;
                int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                string query = "";
                string turno = comboBox4.Text;
                if (Convert.ToString(dataGridView1.CurrentCell.Value) != "P")
                {
                    if (dataGridView1.CurrentCell.Style.BackColor == Color.LightBlue)
                    {
                        query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox1.SelectedIndex == 1)
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                        }
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Pink)
                    {
                        query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox2.SelectedIndex == 0)
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                        }
                        else
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                        }
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Red)
                    {
                        query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox2.SelectedIndex == 0)
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                        }
                        else
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                        }
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Green)
                    {
                        query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox1.SelectedIndex == 1)
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                        }
                    }
                    try
                    {
                        MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                        MySqlCommand cmd = new MySqlCommand(query, cnn);
                        cnn.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                        totalfaltas();
                        faltasinjus();
                        faltasjusti();
                        faltasaño();
                        justiaño();
                        injustiaño();
                        observaciones();
                        //CargarDatos();
                    }
                    catch (Exception el)
                    {
                        MessageBox.Show(el + "");
                    }
                }
            }
        }
        private void eliminarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
                {
                    int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                    int nmes = c_mes.SelectedIndex + 3;
                    int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                    string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                    string turno = comboBox4.Text;
                    string query = "delete from inasistencia where idfalta=" + idfalta(idalumno, fecha, turno);
                    MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                    MySqlCommand cmd = new MySqlCommand(query, cnn);
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                    dataGridView1.CurrentCell.Value = "P";
                    dataGridView1.CurrentCell.Style.BackColor = Color.White;
                    totalfaltas();
                    faltasinjus();
                    faltasjusti();
                    faltasaño();
                    justiaño();
                    injustiaño();
                    observaciones();

                    textBox9.Enabled = false;
                    //CargarDatos();
                }
            }
            catch (Exception el)
            {
                MessageBox.Show(el + "");
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            CargarDatos();
            textBox9.Text = "";
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (Convert.ToString(dataGridView1.CurrentCell.Value) == "P" || Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
                {
                    try
                    {
                        int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                        int nmes = c_mes.SelectedIndex + 3;
                        int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                        string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                        string query;
                        string turno = comboBox4.Text;

                        if (Convert.ToString(dataGridView1.CurrentCell.Value) == "P")
                        {
                            query = "insert into inasistencia (idalumno, fecha, valor, Justif, Turno) values (" + idalumno + ", '" + fecha + "', 1, 'No','" + comboBox4.Text + "')";
                            dataGridView1.CurrentCell.Value = "1";
                            if (comboBox1.Text == "Rojo")
                            {
                                dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                            }

                            textBox9.Enabled = true;
                            button1.Visible = true;
                        }
                        else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1")
                        {
                            query = "Update inasistencia set valor = 0.75 where idfalta =" + idfalta(idalumno, fecha, turno);
                            dataGridView1.CurrentCell.Value = "0,75";

                            textBox9.Enabled = true;
                            button1.Visible = true;
                        }
                        else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75")
                        {
                            query = "Update inasistencia set valor = 0.5 where idfalta =" + idfalta(idalumno, fecha, turno);
                            dataGridView1.CurrentCell.Value = "0,5";

                            textBox9.Enabled = true;
                            button1.Visible = true;
                        }
                        else if (Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5")
                        {
                            query = "Update inasistencia set valor = 0.25 where idfalta =" + idfalta(idalumno, fecha, turno);
                            dataGridView1.CurrentCell.Value = "0,25";

                            button1.Visible = true;
                            textBox9.Enabled = true;
                        }
                        else
                        {
                            query = "delete from inasistencia where idfalta=" + idfalta(idalumno, fecha, turno);
                            dataGridView1.CurrentCell.Value = "P";
                            dataGridView1.CurrentCell.Style.BackColor = Color.White;

                            button1.Visible = false;
                            textBox9.Enabled = false;
                        }
                        MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                        MySqlCommand cmd = new MySqlCommand(query, cnn);
                        cnn.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                        totalfaltas();
                        faltasinjus();
                        faltasjusti();
                        faltasaño();
                        justiaño();
                        injustiaño();
                        observaciones();
                    }
                    catch
                    {
                    }
                }
            }
            else if (e.KeyCode == Keys.F2)
            {
                int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                int nmes = c_mes.SelectedIndex + 3;
                int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                string query = "";
                string turno = comboBox4.Text;
                if (Convert.ToString(dataGridView1.CurrentCell.Value) != "P")
                {
                    if (dataGridView1.CurrentCell.Style.BackColor == Color.LightBlue)
                    {
                        query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox1.SelectedIndex == 1)
                            dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                        else
                            dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Pink)
                    {
                        query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox2.SelectedIndex == 0)
                            dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                        else
                            dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Red)
                    {
                        query = "Update inasistencia set Justif = 'Si' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox2.SelectedIndex == 0)
                            dataGridView1.CurrentCell.Style.BackColor = Color.Green;
                        else
                            dataGridView1.CurrentCell.Style.BackColor = Color.LightBlue;
                    }
                    else if (dataGridView1.CurrentCell.Style.BackColor == Color.Green)
                    {
                        query = "Update inasistencia set Justif = 'No' where idfalta =" + idfalta(idalumno, fecha, turno);
                        if (comboBox1.SelectedIndex == 1)
                            dataGridView1.CurrentCell.Style.BackColor = Color.Red;
                        else
                            dataGridView1.CurrentCell.Style.BackColor = Color.Pink;
                    }
                    try
                    {
                        MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                        MySqlCommand cmd = new MySqlCommand(query, cnn);
                        cnn.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                        totalfaltas();
                        faltasinjus();
                        faltasjusti();
                        faltasaño();
                        justiaño();
                        injustiaño();
                        observaciones();

                    }
                    catch { }
                }
            }
            else if (e.KeyCode == Keys.F9) 
            {
                try
                {
                    if (Convert.ToString(dataGridView1.CurrentCell.Value) == "1" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,75" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,5" || Convert.ToString(dataGridView1.CurrentCell.Value) == "0,25")
                    {
                        int idalumno = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                        int nmes = c_mes.SelectedIndex + 3;
                        int dia = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex) - 1;
                        string fecha = comboBox12.Text + "/" + nmes + "/" + dia;
                        string turno = comboBox4.Text;
                        string query = "delete from inasistencia where idfalta=" + idfalta(idalumno, fecha, turno);
                        MySqlConnection cnn = new MySqlConnection(cadenaconexion());
                        MySqlCommand cmd = new MySqlCommand(query, cnn);
                        cnn.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                        dataGridView1.CurrentCell.Value = "P";
                        dataGridView1.CurrentCell.Style.BackColor = Color.White;
                        totalfaltas();
                        faltasinjus();
                        faltasjusti();
                        faltasaño();
                        justiaño();
                        injustiaño();
                        observaciones();
                        textBox9.Enabled = false;
                        button1.Visible = false;
                        //CargarDatos();
                    }
                }
                catch (Exception el)
                {
                    MessageBox.Show(el + "");
                }

            }
        
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.CurrentCell.Value.ToString() == "1" || dataGridView1.CurrentCell.Value.ToString() == "0.75" || dataGridView1.CurrentCell.Value.ToString() == "0.5" || dataGridView1.CurrentCell.Value.ToString() == "0.25")
                {
                    textBox9.ReadOnly = false;
                    button1.Visible = true;
                    observaciones();
                }
                else
                {
                    button1.Visible = false;
                    textBox9.ReadOnly = true;
                }
                totalfaltas();
                faltasinjus();
                faltasjusti();
                faltasaño();
                justiaño();
                injustiaño();
            }
            catch { }
        }
    }
}

