using Certificados_Clientes_WL.Classes;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Linq.Expressions;
using System.Windows.Forms;

namespace Certificados_Clientes_WL.Repositorios
{
    public class RepositorioCertificados : RepositorioAbstrato<Certificados>
    {
        private ConexaoPG conexaoPG = new ConexaoPG();
        private RepositorioEmpresa repositorioEmpresa = new RepositorioEmpresa();
        private RepositorioUsuario repositorioUsuario = new RepositorioUsuario();
        public List<Certificados> Certificados{ get; set; }
        public int QuantidadeCertificados { get; set; }
        public override void Add(Certificados x)
        {
            throw new NotImplementedException();
        }

        public override IEnumerable<Certificados> Get(Expression<Func<Certificados, bool>> predicate)
        {
            var certificados = (List<Certificados>)GetAll();
            Certificados = certificados.AsQueryable().Where(predicate).ToList();
            QuantidadeCertificados = Certificados.Count;
            return certificados.AsQueryable().Where(predicate).ToList();
        }

        public override IEnumerable<Certificados> GetAll()
        {
            string listaCertificados = @"SELECT * FROM CERTIFICADOS
                                            INNER JOIN PRESTACAOSERVICO ON CERTIFICADOS.IDEMPRESA = PRESTACAOSERVICO.IDEMPRESA
                                            WHERE PRESTACAOSERVICO.STATUS NOT IN('Baixa', 'Extinta', 'Suspensa', 'Retirada/Dispensada') AND VALIDO = 'true'
                                            ORDER BY CERTIFICADOS.DATAVENCIMENTO";

            NpgsqlConnection connect = conexaoPG.ConexaoBanco();

            try
            {
                connect.Open();
                var cmd = connect.CreateCommand();
                cmd.CommandText = listaCertificados;

                var cmdDt = new NpgsqlDataAdapter(cmd);
                var dataTable = new DataTable();
                cmdDt.Fill(dataTable);

                List<Certificados> certificados = new List<Certificados>();

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    Certificados certificado = new Certificados();

                    certificado.Id = (int)dataTable.Rows[i][0];
                    certificado.NomeCertificado = dataTable.Rows[i][1].ToString();
                    certificado.DataDeVencimento = (DateTime)dataTable.Rows[i][3];
                    certificado.Valido = certificado.DataDeVencimento >= DateTime.Now? "Ativo" : "Vencido";
                    certificado.DataCadastro = (DateTime)dataTable.Rows[i][5];
                    certificado.Senha = dataTable.Rows[i][7].ToString();
                    certificado.Empresa = repositorioEmpresa.Get(x => x.Id.Equals((int)dataTable.Rows[i][8])); 
                    certificado.Tipo = dataTable.Rows[i][9].ToString();
                    certificado.Certificadora = dataTable.Rows[i][13].ToString();

                    certificados.Add(certificado);
                }
                QuantidadeCertificados = certificados.Count;
                Certificados = certificados;
                connect.Close();
                return Certificados;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            finally
            {
                connect.Close();
            }
        }

        public override void Remove(Certificados x) => throw new NotImplementedException();

        public override void Update(Certificados x) => throw new NotImplementedException();

        public DataTable GetTipoCertificado()
        {
            NpgsqlConnection connect = conexaoPG.ConexaoBanco();
            DataTable dataTable = new DataTable();
            
            connect.Open();

            string consultaTipoCertificado = @"SELECT DISTINCT  (tipo) as tipocertificado
                                                    FROM certificados 
                                                    ORDER BY 1"
            ;

            NpgsqlDataAdapter Adpt = new NpgsqlDataAdapter(consultaTipoCertificado, connect);
            Adpt.Fill(dataTable);
            connect.Close();
            return dataTable;
        }
    }
}
