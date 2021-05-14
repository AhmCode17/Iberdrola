using System;
using System.Data.OleDb;

namespace ResumenesIBerdrola.Data
{
    public class MsAccessDataContext
    {

        public Result SaveCentralTotal(ResumenModel model)
        {
            Result result = new Result();
            try
            {
                string connstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Babel\Documents\Iberdrola\ResumenDb.accdb";

                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "INSERT INTO CentralTotal (EnergiaBase, EnergiaIntermedia, EnergiaPunta,EnergiaSemiPunta,EnergiaTotal,DemandaBase,DemandaIntermedia,DemandaPunta,DemandaSemiPunta,EnergiaReactiva,Fp,Periodo,FkConcepto,FkCentral) " +
                        "VALUES(@EnergiaBase, @EnergiaIntermedia, @EnergiaPunta,@EnergiaSemiPunta,@EnergiaTotal,@DemandaBase,@DemandaIntermedia,@DemandaPunta,@DemandaSemiPunta,@EnergiaReactiva,@Fp,@Periodo,@FkConcepto,@FkCentral);";

                    OleDbCommand comando = new OleDbCommand(sql, con);
                    comando.Parameters.AddWithValue("@EnergiaBase", model.KwBase);
                    comando.Parameters.AddWithValue("@EnergiaIntermedia",22);
                    comando.Parameters.AddWithValue("@EnergiaPunta", 1.3);
                    comando.Parameters.AddWithValue("@EnergiaSemiPunta", 1.5);
                    comando.Parameters.AddWithValue("@EnergiaTotal", 1.6);
                    comando.Parameters.AddWithValue("@DemandaBase", 1.7);
                    comando.Parameters.AddWithValue("@DemandaIntermedia", 1.8);
                    comando.Parameters.AddWithValue("@DemandaPunta", 1.9);
                    comando.Parameters.AddWithValue("@DemandaSemiPunta", 1.10);
                    comando.Parameters.AddWithValue("@EnergiaReactiva", 1.11);
                    comando.Parameters.AddWithValue("@Fp", 1.12);
                    comando.Parameters.AddWithValue("@Periodo", "Marzo 2021");
                    comando.Parameters.AddWithValue("@FkConcepto", 1);
                    comando.Parameters.AddWithValue("@FkCentral", 1);
                    comando.ExecuteNonQuery();

                    con.Close();
                }
                result.Success = true;
                result.Msg = "Se proceso el archivo";

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }

            return result;
        }


        public Result SavePuntoDeCarga()
        {
            Result result = new Result();
            try
            {
                string connstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Babel\Documents\Iberdrola\ResumenDb.accdb";

                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "INSERT INTO CentralTotal (EnergiaBase, EnergiaIntermedia, EnergiaPunta,EnergiaSemiPunta,EnergiaTotal,DemandaBase,DemandaIntermedia,DemandaPunta,DemandaSemiPunta,EnergiaReactiva,Fp,Periodo,FkConcepto,FkCentral) " +
                        "VALUES(@EnergiaBase, @EnergiaIntermedia, @EnergiaPunta,@EnergiaSemiPunta,@EnergiaTotal,@DemandaBase,@DemandaIntermedia,@DemandaPunta,@DemandaSemiPunta,@EnergiaReactiva,@Fp,@Periodo,@FkConcepto,@FkCentral);";

                    OleDbCommand comando = new OleDbCommand(sql, con);
                    comando.Parameters.AddWithValue("@EnergiaBase", 1.1);
                    comando.Parameters.AddWithValue("@EnergiaIntermedia", 1.2);
                    comando.Parameters.AddWithValue("@EnergiaPunta", 1.3);
                    comando.Parameters.AddWithValue("@EnergiaSemiPunta", 1.5);
                    comando.Parameters.AddWithValue("@EnergiaTotal", 1.6);
                    comando.Parameters.AddWithValue("@DemandaBase", 1.7);
                    comando.Parameters.AddWithValue("@DemandaIntermedia", 1.8);
                    comando.Parameters.AddWithValue("@DemandaPunta", 1.9);
                    comando.Parameters.AddWithValue("@DemandaSemiPunta", 1.10);
                    comando.Parameters.AddWithValue("@EnergiaReactiva", 1.11);
                    comando.Parameters.AddWithValue("@Fp", 1.12);
                    comando.Parameters.AddWithValue("@Periodo", "Marzo 2021");
                    comando.Parameters.AddWithValue("@FkConcepto", 1);
                    comando.Parameters.AddWithValue("@FkCentral", 1);
                    comando.ExecuteNonQuery();

                    con.Close();
                }
                result.Success = true;
                result.Msg = "Se proceso el archivo";

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }
            return result;
        }

    }
}
