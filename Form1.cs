using LinqToExcel;
using MySql.Data.MySqlClient;
using PROYECTO_BD.Model;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROYECTO_BD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonCargar_Click(object sender, EventArgs e)
        {
            //leemos el archivo con sldocument
            SLDocument sl = new SLDocument(@"C:\Users\13237\OneDrive\Desktop\DatosDB.xlsx"); //ubicacion
            SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics(); //traer las propiedades del archivo

            int ultimaFila = propiedades.EndRowIndex; //saber cuantas filas existen, trae la ultima fila
            MySqlConnection conexionBD = ClsSQL.conexion(); 
            conexionBD.Open();//se abre la conexion

            for (int x = 2; x <= ultimaFila; x++)//leer todas las filas
            {
                //se trae el codigo
                string sql = "INSERT INTO tabla_alumnos(Correlativo, Nombre, Parcial1, Parcial2, Parcial3) " +
                   "VALUES(@Correlativo, @Nombre, @Parcial1, @Parcial2, @Parcial3)"; 

                try
                {
                    //transaccion a mysql
                    //con el objeto se agrega el alias, el valor que se le asigna, la columna 
                    MySqlCommand comando = new MySqlCommand(sql, conexionBD);
                    comando.Parameters.AddWithValue("@Correlativo", sl.GetCellValueAsString("A" + x));
                    comando.Parameters.AddWithValue("@Nombre", sl.GetCellValueAsString("B" + x));
                    comando.Parameters.AddWithValue("@Parcial1", sl.GetCellValueAsString("C" + x));
                    comando.Parameters.AddWithValue("@Parcial2", sl.GetCellValueAsString("D" + x));
                    comando.Parameters.AddWithValue("@Parcial3", sl.GetCellValueAsString("E" + x));
                    comando.ExecuteNonQuery(); //se ejecuta la insercion
                }//si hay un error se utiliza mysqlexception
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                } 
            }
            MessageBox.Show("El archivo se ha cargado");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument(@"C:\Users\13237\OneDrive\Desktop\DatosDB.xlsx");
            SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics();

            int ultimaFila = propiedades.EndRowIndex;
            SqlConnection conexionBD = ClsSQLServer.conexionServer();
            conexionBD.Open();

            for (int x = 2; x <= ultimaFila; x++)//leer todas las filas
            {
                string sqlServer = "INSERT INTO tb_alumnos(Correlativo, Nombre, Parcial1, Parcial2, Parcial3) " +
                   "VALUES(@Correlativo, @Nombre, @Parcial1, @Parcial2, @Parcial3)";

                try
                {
                    SqlCommand comando = new SqlCommand(sqlServer, conexionBD);
                    comando.Parameters.AddWithValue("@Correlativo", sl.GetCellValueAsString("A" + x));
                    comando.Parameters.AddWithValue("@Nombre", sl.GetCellValueAsString("B" + x));
                    comando.Parameters.AddWithValue("@Parcial1", sl.GetCellValueAsString("C" + x));
                    comando.Parameters.AddWithValue("@Parcial2", sl.GetCellValueAsString("D" + x));
                    comando.Parameters.AddWithValue("@Parcial3", sl.GetCellValueAsString("E" + x));
                    comando.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            MessageBox.Show("El archivo se ha cargado");
        }

        private void buttonImportar_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string strRuta = openFileDialog1.FileName;
            ExcelQueryFactory excelQueryFactory = new ExcelQueryFactory(strRuta);
            var oDatos = (from ss in excelQueryFactory.Worksheet("DatosDB")
                          let item = new hojaExcel
                          {
                              Correlativo = ss[1].Cast<string>(),
                              Nombre = ss[1].Cast<string>(),
                              Parcial1 = ss[1].Cast<string>(),
                              Parcial2 = ss[1].Cast<string>(),
                              Parcial3 = ss[1].Cast<string>()
                          }
                          select item).ToList();
            importarExcel(oDatos);
        }
        //crear un destructor
        private void importarExcel(List<hojaExcel> oDatos)
        {
            
            using (var objContex = new db_alumnosEntities())
            {
                
                foreach (var item in oDatos)
                {
                    if (item.Correlativo != null && item.Nombre !=null && item.Parcial1 !=null && item.Parcial2 !=null && item.Parcial3 != null)
                    {
                        
                    }
                }
            }
        }

        public class hojaExcel
        {
            public string Correlativo { get; set; }
            public string Nombre { get; set; }
            public string Parcial1 { get; set; }
            public string Parcial2 { get; set; }
            public string Parcial3 { get; set; }
        }


        //to read the excel data into datagridview
        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection laConexion = new OleDbConnection(@"C: \Users\13237\OneDrive\Desktop\");
            laConexion.Open();
            OleDbDataAdapter elAd = new OleDbDataAdapter("Select * from[sheet1$]", laConexion);
            DataSet SD = new DataSet();
            DataTable DT = new DataTable();
            elAd.Fill(DT);
            this.dataGridView1.DataSource = DT.DefaultView;
        }
    }
}
