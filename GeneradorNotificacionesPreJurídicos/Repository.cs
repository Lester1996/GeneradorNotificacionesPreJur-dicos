using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using System.Threading.Tasks;
using GeneradorNotificacionesPreJurídicos.Models;
using System.Configuration;



namespace GeneradorNotificacionesPreJurídicos
{
    public class Repository
    {


        string connectionString = "Data Source=(DESCRIPTION =(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST = 192.168.100.89)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = SGCPRO)));User ID=CJM01521;Password=CJM01521";
        string query = "";

        public List<InformacionCliente> GetInformacionClientes(List<string> claves)
        {
            string clavesSeparadasPorComa = string.Join(",", claves);

            try
            {
                List<InformacionCliente> resultadoConsultaList = new List<InformacionCliente>();

                string query = $@"SELECT PAN_CUENTA_2 CLAVE,
                            A.NIS_RAD,
                            REPLACE(NOMBRE_CLIENTE, '   ', ' ') AS NOMBRE_CLIENTE,
                            REPLACE(REF_DIR, '      ', ' ') AS REF_DIR,
                            TO_CHAR(F_ULT_PAGO, 'YYYY-MM-DD') ULTIMO_PAGO,
                            SYSDATE AS FECHA_ACTUAL,
                            INITCAP(B.NOM_AREA) AS NOM_AREA,
                            TO_NUMBER(TO_CHAR(SYSDATE, 'DD')) AS DIA_ACTUAL,
                            TO_NUMBER(TO_CHAR(SYSDATE, 'MM')) AS MES_ACTUAL,
                            TO_NUMBER(TO_CHAR(SYSDATE, 'YYYY')) AS ANIO_ACTUAL,
                            REPLACE(LOWER(TO_CHAR(SYSDATE, 'MONTH', 'NLS_DATE_LANGUAGE=SPANISH')), ' ', '') AS MES_ACTUAL_TEXTO,
                            INITCAP(TO_CHAR(SYSDATE, 'DAY', 'NLS_DATE_LANGUAGE=SPANISH')) AS DIA_ACTUAL_TEXTO,
                            D.DEUDA_TOTAL
                     FROM MAESTRO_SUSCRIPTORES A
                     INNER JOIN AREAS B
                     ON A.COD_AREA = B.COD_AREA
                            INNER JOIN DEUDA D
                            ON A.NIS_RAD=D.NIS_RAD
                     WHERE A.PAN_CUENTA_2 IN ({clavesSeparadasPorComa})";

                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            foreach (DataRow row in dataTable.Rows)
                            {
                                resultadoConsultaList.Add(new InformacionCliente
                                {
                                    Clave = Convert.ToString(row["CLAVE"]),
                                    NisRad = Convert.ToInt32(row["NIS_RAD"]),
                                    NombreCliente = row["NOMBRE_CLIENTE"].ToString(),
                                    DireccionCliente = row["REF_DIR"].ToString(),
                                    FechaUltimoPago = DateTime.Parse(row["ULTIMO_PAGO"].ToString()),
                                    FechaActual = DateTime.Parse(row["FECHA_ACTUAL"].ToString()),
                                    DiaActual = Convert.ToInt32(row["DIA_ACTUAL"]),
                                    MesActual = Convert.ToInt32(row["MES_ACTUAL"]),
                                    AnioActual = Convert.ToInt32(row["ANIO_ACTUAL"]),
                                    Area = row["NOM_AREA"].ToString(),
                                    MesActualText = row["MES_ACTUAL_TEXTO"].ToString(),
                                    DiaActualText = row["DIA_ACTUAL_TEXTO"].ToString(),
                                    DeudaTotal   = double.Parse(row["DEUDA_TOTAL"].ToString())
                                });
                            }
                        }
                    }
                }
                return resultadoConsultaList;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public List<SaldoFinanciado> GetSaldoFinanciados(List<string> nisRads)
        {
            try
            {
                List<SaldoFinanciado> resultadoConsultaList = new List<SaldoFinanciado>();
                string nisRadsSeparadasPorComa = string.Join(",", nisRads);

                string query = $@"SELECT M.NIS_RAD, A.SALDO_FINANCIADO
                          FROM MACUERDOS M
                          INNER JOIN (
                              SELECT C.NIS_RAD NIS_RAD, COUNT(C.NIS_RAD) CUOTAS_PENDIENTES_DE_FACTURAR, 
                                     SUM(C.IMP_AM_CUOTA + C.IMP_INT_CUOTA) SALDO_FINANCIADO
                              FROM CUOTAS_PL C
                              LEFT JOIN ESTADOS E ON C.EST_CUOTA = E.ESTADO
                              WHERE DESC_EST IN ('Pendiente de Facturar')
                              HAVING SUM(C.IMP_AM_CUOTA + C.IMP_INT_CUOTA) > 0
                              GROUP BY C.NIS_RAD
                          ) A ON M.NIS_RAD = A.NIS_RAD
                          WHERE M.EST_ACU = 'EA001'
                          AND M.NIS_RAD IN ({nisRadsSeparadasPorComa})";

                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            if (dataTable.Rows.Count == 0)
                            {
                                // Si no hay filas en el DataTable, añadir un objeto con SaldoF = 0
                                resultadoConsultaList.Add(new SaldoFinanciado
                                {
                                    Nisrad = 0,
                                    SaldoF = 0,
                                });
                            }
                            else
                            {
                                foreach (DataRow row in dataTable.Rows)
                                {
                                    decimal saldoFinanciado;
                                    if (!decimal.TryParse(row["SALDO_FINANCIADO"].ToString(), out saldoFinanciado))
                                    {
                                        saldoFinanciado = 0;
                                    }

                                    resultadoConsultaList.Add(new SaldoFinanciado
                                    {
                                        Nisrad = int.Parse(row["NIS_RAD"].ToString()),
                                        SaldoF = saldoFinanciado,
                                    });
                                }
                            }
                        }
                    }
                }

                return resultadoConsultaList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        public List<DeudaCliente> GetDeudaClientes(int nisRad)
        {

            try
            {
                List<DeudaCliente> resultadoConsultaList = new List<DeudaCliente>();

                query = $@"SELECT NIS_RAD, TO_CHAR(DEUDA_TOTAL, 'FM999,999,999.99') AS DEUDA_TOTAL FROM DEUDA
                        WHERE NIS_RAD = {nisRad}";

                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            foreach (DataRow row in dataTable.Rows)
                            {
                                resultadoConsultaList.Add(new DeudaCliente
                                {
                                    NisRad = int.Parse(row["NIS_RAD"].ToString()),
                                    Deuda = decimal.Parse(row["DEUDA_TOTAL"].ToString()),
                                });
                            }
                        }
                    }
                }

                return resultadoConsultaList;



            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }



        public List<InfoLider> GetInfoLiders(List<string> areas)
        {
            string cadenaConexion = @"Data Source=192.168.100.28;Initial Catalog=CobroRecaudo;User ID=appPortalCobroRec;Password=AppPC0bRe32019$";
            try
            {
                List<InfoLider> resultadoConsultaList = new List<InfoLider>();
                string areasSeparadasPorComa = string.Join(",", areas.Select(a => $"'{a.ToUpper()}'"));

                string query = $@"SELECT LOWER(SECTOR) AS SECTOR, LIDER_COBRO, ANALISTA_COBRO_PREJURIDICO, TELEFONO, EMAIL ,FIRMA_ELECTRONICA
                             FROM INFO_COBRO_PREJURIDICO
                             WHERE SECTOR IN ({areasSeparadasPorComa})";

                using (SqlConnection connection = new SqlConnection(cadenaConexion))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            foreach (DataRow row in dataTable.Rows)
                            {
                                resultadoConsultaList.Add(new InfoLider
                                {
                                    SectorL = row["SECTOR"].ToString(),
                                    NombreL = row["LIDER_COBRO"].ToString(),
                                    AnalistaC = row["ANALISTA_COBRO_PREJURIDICO"].ToString(),
                                    Telefono = row["TELEFONO"].ToString(),
                                    Email = row["EMAIL"].ToString(),
                                    FirmaElectronica = (byte[])row["FIRMA_ELECTRONICA"]
                                });
                            }
                        }
                    }
                }
                return resultadoConsultaList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }



        public List<InfoGestor> GetInfoGestor(List<string> Gestores)
        {
            string cadenaConexion = @"Data Source=192.168.100.28;Initial Catalog=CobroRecaudo;User ID=appPortalCobroRec;Password=AppPC0bRe32019$";
            try
            {
                List<InfoGestor> resultadoConsultaList = new List<InfoGestor>();
                string GestoresSeparadasPorComa = string.Join(",", Gestores.Select(a => $"'{a.ToUpper()}'"));

                string query = $@"SELECT LOWER(SECTOR) AS SECTOR, TELEFONO, NOMBRE, USUARIO, CARGO
                             FROM INFO_GESTOR_PREJURIDICO
                             WHERE TRIM(USUARIO) IN ({GestoresSeparadasPorComa})";

                using (SqlConnection connection = new SqlConnection(cadenaConexion))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            foreach (DataRow row in dataTable.Rows)
                            {
                                resultadoConsultaList.Add(new InfoGestor
                                {
                                    SectorG = row["SECTOR"].ToString(),
                                    NombreG = row["NOMBRE"].ToString(),
                                    Telefono = row["TELEFONO"].ToString(),
                                    Usuario = row["USUARIO"].ToString(),
                                    Cargo = row["CARGO"].ToString(),

                                });
                            }
                        }
                    }
                }
                return resultadoConsultaList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }




}




