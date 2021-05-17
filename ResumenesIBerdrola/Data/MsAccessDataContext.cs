using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace ResumenesIBerdrola.Data
{
    public class MsAccessDataContext
    {
        string connstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Babel\Documents\Iberdrola\ResumenDb.accdb";

        public Result SaveResumen(ResumenBaseModel model)
        {
            Result result = new Result();
            try
            {
                var existResumen = (ResumenBaseModel)GetResumen(model).Data;
                if (existResumen.Id > 0)
                {
                    result.Success = false;
                    result.Msg = string.Format("Ya existe una central: {0} con periodo periodo {0}", model.Central, model.Periodo);
                }
                else
                {
                    using (OleDbConnection con = new OleDbConnection(connstring))
                    {
                        con.Open();

                        string sql = "INSERT INTO Resumen (FkCentral, Periodo, FechaCreacion) " +
                            "VALUES(@FkCentral, @Periodo, @FechaCreacion);";

                        OleDbCommand comando = new OleDbCommand(sql, con);
                        comando.Parameters.AddWithValue("@FkCentral", model.FkCentral);
                        comando.Parameters.AddWithValue("@Periodo", model.Periodo);
                        comando.Parameters.AddWithValue("@FechaCreacion", model.FechaCreacion.ToString());
                        comando.ExecuteNonQuery();

                        con.Close();
                    }
                    result.Data = GetResumen(model).Data;
                    result.Success = true;
                    result.Msg = "Se proceso el archivo";
                }

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }

            return result;
        }

        public Result GetResumen(ResumenBaseModel model)
        {
            Result result = new Result();
            OleDbDataReader reader = null;
            try
            {
                ResumenBaseModel resumen = new ResumenBaseModel();
                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "SELECT * FROM Resumen WHERE FkCentral=@fkCentral AND Periodo=@periodo";
                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    cmd.Parameters.AddWithValue("@fkCentral", model.FkCentral);
                    cmd.Parameters.AddWithValue("@periodo", model.Periodo);
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        resumen = new ResumenBaseModel
                        {
                            Id = reader.GetInt32(reader.GetOrdinal("Id")),
                            FkCentral = reader.GetInt32(reader.GetOrdinal("FkCentral")),
                            Periodo = reader.GetString(reader.GetOrdinal("Periodo"))
                        };

                    }
                    con.Close();
                }
                result.Success = true;
                result.Data = resumen;
                result.Msg = "Consulta existosa";

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }
            finally
            {
                if (reader != null) reader.Close();
            }
            return result;
        }
        public Result SaveCentralTotal(ResumenModel model)
        {
            Result result = new Result();
            try
            {

                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "INSERT INTO CentralTotal (EnergiaBase, EnergiaIntermedia, EnergiaPunta,EnergiaSemiPunta,EnergiaTotal,DemandaBase,DemandaIntermedia,DemandaPunta,DemandaSemiPunta,EnergiaReactiva,Fp,Periodo,FkConcepto,FkResumen,Concepto) " +
                        "VALUES(@EnergiaBase, @EnergiaIntermedia, @EnergiaPunta,@EnergiaSemiPunta,@EnergiaTotal,@DemandaBase,@DemandaIntermedia,@DemandaPunta,@DemandaSemiPunta,@EnergiaReactiva,@Fp,@Periodo,@FkConcepto,@FkResumen,@Concepto);";

                    OleDbCommand comando = new OleDbCommand(sql, con);
                    comando.Parameters.AddWithValue("@EnergiaBase", model.KwhBase);
                    comando.Parameters.AddWithValue("@EnergiaIntermedia", model.KwhIntermedia);
                    comando.Parameters.AddWithValue("@EnergiaPunta", model.KwhPunta);
                    comando.Parameters.AddWithValue("@EnergiaSemiPunta", model.KwhSemiPunta);
                    comando.Parameters.AddWithValue("@EnergiaTotal", model.KwhTotales);
                    comando.Parameters.AddWithValue("@DemandaBase", model.KwBase);
                    comando.Parameters.AddWithValue("@DemandaIntermedia", model.KwIntermedia);
                    comando.Parameters.AddWithValue("@DemandaPunta", model.KwPunta);
                    comando.Parameters.AddWithValue("@DemandaSemiPunta", model.KwSemiPunta);
                    comando.Parameters.AddWithValue("@EnergiaReactiva", model.KwKvarh);
                    comando.Parameters.AddWithValue("@Fp", model.KwFp);
                    comando.Parameters.AddWithValue("@Periodo", model.Periodo);
                    comando.Parameters.AddWithValue("@FkConcepto", model.FkConcepto);
                    comando.Parameters.AddWithValue("@FkResumen", model.FkResumen);
                    comando.Parameters.AddWithValue("@Concepto", model.Concepto);
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

        public Result GetConcepto()
        {
            Result result = new Result();
            OleDbDataReader reader = null;
            try
            {
                List<ConceptoModel> lstConcepto = new List<ConceptoModel>();
                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "SELECT * FROM Concepto";
                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        var concepto = new ConceptoModel
                        {
                            Id = reader.GetInt32(reader.GetOrdinal("Id")),
                            Concepto = reader.GetString(reader.GetOrdinal("Concepto")),
                            TipoConcepto = reader.GetString(reader.GetOrdinal("TipoConcepto"))
                        };
                        lstConcepto.Add(concepto);
                    }
                    con.Close();
                }
                result.Success = true;
                result.Data = lstConcepto;
                result.Msg = "Consulta existosa";

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }
            finally
            {
                if (reader != null) reader.Close();
            }

            return result;
        }

        public Result GetCentral()
        {
            Result result = new Result();
            OleDbDataReader reader = null;
            try
            {
                List<CentralModel> lstCentral = new List<CentralModel>();
                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "SELECT * FROM Central";

                    OleDbCommand cmd = new OleDbCommand(sql, con);
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        var central = new CentralModel
                        {
                            Id = reader.GetInt32(reader.GetOrdinal("Id")),
                            RazonSocial = reader.GetString(reader.GetOrdinal("RazonSocial")),
                            Cliente = reader.GetString(reader.GetOrdinal("Cliente"))
                        };
                        lstCentral.Add(central);
                    }

                    con.Close();
                }
                result.Success = true;
                result.Data = lstCentral;
                result.Msg = "Se proceso el archivo";

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Msg = ex.Message;
            }

            return result;
        }
        public Result SavePuntoDeCarga(ResumenModel model)
        {
            Result result = new Result();
            try
            {

                using (OleDbConnection con = new OleDbConnection(connstring))
                {
                    con.Open();

                    string sql = "INSERT INTO PuntodeCarga (EnergiaBase, EnergiaIntermedia, EnergiaPunta,EnergiaSemiPunta,EnergiaTotal,DemandaBase,DemandaIntermedia,DemandaPunta,DemandaTotal,Fp,CapacidadTotal,DistribucionTotal,Descripcion,Rpu,NombreCliente,FkConcepto,FkResumen) " +
                        "VALUES(@EnergiaBase, @EnergiaIntermedia, @EnergiaPunta,@EnergiaSemiPunta,@EnergiaTotal,@DemandaBase,@DemandaIntermedia,@DemandaPunta,@DemandaTotal,@Fp,@CapacidadTotal,@DistribucionTotal,@Descripcion,@Rpu,@NombreCliente,@FkConcepto,@FkResumen);";

                    OleDbCommand comando = new OleDbCommand(sql, con);
                    comando.Parameters.AddWithValue("@EnergiaBase", model.KwhBase);
                    comando.Parameters.AddWithValue("@EnergiaIntermedia", model.KwhIntermedia);
                    comando.Parameters.AddWithValue("@EnergiaPunta", model.KwhPunta);
                    comando.Parameters.AddWithValue("@EnergiaSemiPunta", model.KwhSemiPunta);
                    comando.Parameters.AddWithValue("@EnergiaTotal", model.KwhTotales);
                    comando.Parameters.AddWithValue("@DemandaBase", model.KwBase);
                    comando.Parameters.AddWithValue("@DemandaIntermedia", model.KwIntermedia);
                    comando.Parameters.AddWithValue("@DemandaPunta", model.KwPunta);
                    comando.Parameters.AddWithValue("@DemandaTotal", model.KwSemiPunta);
                    comando.Parameters.AddWithValue("@Fp", model.KwFp);

                    comando.Parameters.AddWithValue("@CapacidadTotal", model.CapacidadTotal);
                    comando.Parameters.AddWithValue("@DistribucionTotal", model.DistribucionTotal);

                    comando.Parameters.AddWithValue("@Descripcion", model.Descripcion);
                    comando.Parameters.AddWithValue("@Rpu", model.Rpu);
                    comando.Parameters.AddWithValue("@NombreCliente", model.NombreCliente);

                    comando.Parameters.AddWithValue("@FkConcepto", model.FkConcepto);
                    comando.Parameters.AddWithValue("@FkResumen", model.FkResumen);
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
