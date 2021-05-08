using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROYECTO_BD
{
    class ClsSQLServer
    {

        public static SqlConnection conexionServer() 
        {
           
            string conexionS = @"Data source=(Localdb)\\Darlin; Inicial Catalog=db_alumnos; Integrated Security=True";

            try
            {
                SqlConnection conexionBD = new SqlConnection(conexionS);
                return conexionBD;
            }
            catch (SqlException ex)
            {
                Console.WriteLine("Error:" + ex.Message);
                return null;
            }

        }
    }
}
