using Certificados_Clientes_WL.Classes;
using Certificados_Clientes_WL.Repositorios;
using Certificados_Clientes_WL.Resources;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Fonte = iTextSharp.text.Font;
namespace Certificados_Clientes_WL.Formularios
{
    public partial class FrmCertificados : Form
    {
        RepositorioCertificados repositorioCertificados = new RepositorioCertificados();
        readonly DataTable dtTable = new DataTable();
        public int TotalPaginas = 1;
        public FrmCertificados()
        {
            InitializeComponent();
            CarregaTipoCertificado();
            AddColunasTabela();
            comboBoxStatus.SelectedIndex = 0;
            comboBoxMes.SelectedIndex = 0;
        }

        private void Carregar(object sender, EventArgs e)
        {
            string status = comboBoxStatus.Text;
            string mesEscrito = comboBoxMes.Text;
            string tipo = comboBoxTipo.Text;
            bool filtroStatus = !status.Equals("<Todos>");
            bool filtroMesEscrito = !mesEscrito.Equals("<Todos>");
            bool filtroTipo = !tipo.Equals("<Todos>");
            string mes = "";
            string anoAnalise = DateTime.Now.Year.ToString();
            if (filtroMesEscrito)
            {
                if (mesEscrito.Equals("JANEIRO"))
                {
                    mes = "01";
                }
                else if (mesEscrito.Equals("FEVEREIRO"))
                {
                    mes = "02";
                }
                else if (mesEscrito.Equals("MARÇO"))
                {
                    mes = "03";
                }
                else if (mesEscrito.Equals("ABRIL"))
                {
                    mes = "04";
                }
                else if (mesEscrito.Equals("MAIO"))
                {
                    mes = "05";
                }
                else if (mesEscrito.Equals("JUNHO"))
                {
                    mes = "06";
                }
                else if (mesEscrito.Equals("JULHO"))
                {
                    mes = "07";
                }
                else if (mesEscrito.Equals("AGOSTO"))
                {
                    mes = "08";
                }
                else if (mesEscrito.Equals("SETEMBRO"))
                {
                    mes = "09";
                }
                else if (mesEscrito.Equals("OUTUBRO"))
                {
                    mes = "10";
                }
                else if (mesEscrito.Equals("NOVEMBRO"))
                {
                    mes = "11";
                }
                else if (mesEscrito.Equals("DEZEMBRO"))
                {
                    mes = "12";
                }
            }

            if (!status.Equals("<Todos>"))
            {
                status = status.Equals("Valido") ? "true" : "false";
            }

            if (filtroStatus && filtroMesEscrito && filtroTipo)
            {
                var data = new DateTime(DateTime.Now.Year, Convert.ToInt32(mes), 01);
                var ultimoDia = DateTime.DaysInMonth(data.Year, data.Month);
                var dataInicio = anoAnalise + '/' + mes + '/' + "01";
                var dataFinal = anoAnalise + '/' + mes + '/' + ultimoDia;

                if (status.Equals("true"))
                {
                    GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo) && x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal) && x.DataDeVencimento >= DateTime.Now));
                }
                else
                {
                    GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo) && x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal) && x.DataDeVencimento < DateTime.Now));
                }
            }
            else if (filtroStatus && filtroMesEscrito && !filtroTipo)
            {
                var data = new DateTime(DateTime.Now.Year, Convert.ToInt32(mes), 01);
                var ultimoDia = DateTime.DaysInMonth(data.Year, data.Month);
                var dataInicio = anoAnalise + '/' + mes + '/' + "01";
                var dataFinal = anoAnalise + '/' + mes + '/' + ultimoDia;

                if (status.Equals("true"))
                {
                    GridCertificados(repositorioCertificados.Get(x => x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal) && x.DataDeVencimento >= DateTime.Now));
                }
                else
                {
                    GridCertificados(repositorioCertificados.Get(x => x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal) && x.DataDeVencimento < DateTime.Now));
                }
            }
            else if (filtroStatus && !filtroMesEscrito && filtroTipo)
            { 
                if (status.Equals("true"))
                {
                    GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo) && 
                    x.DataDeVencimento >= DateTime.Now));
                }
                else
                {
                    GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo) &&
                    x.DataDeVencimento < DateTime.Now));
                }
            }
            else if(!filtroStatus && filtroMesEscrito && filtroTipo)
            {
                var data = new DateTime(DateTime.Now.Year, Convert.ToInt32(mes), 01);
                var ultimoDia = DateTime.DaysInMonth(data.Year, data.Month);
                var dataInicio = anoAnalise + '/' + mes + '/' + "01";
                var dataFinal = anoAnalise + '/' + mes + '/' + ultimoDia;

                GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo) && x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal)));
            }
            else if(filtroStatus && !filtroMesEscrito && !filtroTipo)
            {
                if (status.Equals("true"))
                {
                    GridCertificados(repositorioCertificados.Get(x => x.DataDeVencimento >= DateTime.Now));
                }
                else
                {
                    GridCertificados(repositorioCertificados.Get(x => x.DataDeVencimento < DateTime.Now));
                }
            }
            else if(!filtroStatus && filtroMesEscrito && !filtroTipo)
            {
                var data = new DateTime(DateTime.Now.Year, Convert.ToInt32(mes), 01);
                var ultimoDia = DateTime.DaysInMonth(data.Year, data.Month);
                var dataInicio = anoAnalise + '/' + mes + '/' + "01";
                var dataFinal = anoAnalise + '/' + mes + '/' + ultimoDia;

                GridCertificados(repositorioCertificados.Get(x => x.DataDeVencimento >= Convert.ToDateTime(dataInicio)
                        && x.DataDeVencimento <= Convert.ToDateTime(dataFinal)));
            }
            else if(!filtroStatus && !filtroMesEscrito && filtroTipo)
            {
                GridCertificados(repositorioCertificados.Get(x => x.Tipo.Equals(tipo)));
            }
            else
            {
                GridCertificados(repositorioCertificados.GetAll());
            }
            label4.Text = repositorioCertificados.QuantidadeCertificados.ToString();
        }
       
        private void AddColunasTabela()
        {
            dtTable.Columns.Add("Cód. Empresa", typeof(int));
            dtTable.Columns.Add("Filial", typeof(int));
            dtTable.Columns.Add("Nome Empresa", typeof(string));
            dtTable.Columns.Add("Valido", typeof(string));
            dtTable.Columns.Add("Tipo", typeof(string));
            dtTable.Columns.Add("Certificadora", typeof(string));
            dtTable.Columns.Add("Data Vencimento", typeof(DateTime));

        }

        private void CarregaTipoCertificado()
        {
            DataTable objDataTable = repositorioCertificados.GetTipoCertificado();
            comboBoxTipo.Items.Add("<Todos>");

            foreach (DataRow dataRow in objDataTable.Rows)
            {
                string tipocertificado = dataRow["Tipocertificado"].ToString();
                comboBoxTipo.Items.Add(tipocertificado);
            }
            comboBoxTipo.SelectedIndex = 0;
        }

        private void GridCertificados(IEnumerable<Certificados> certificados)
        {
            DadosDataTable(certificados);
            bindingSource1.DataSource = dtTable;
            dgvCertificados.DataSource = bindingSource1;

            dgvCertificados.Columns[0].Width = 50;
            dgvCertificados.Columns[1].Width = 30;
            dgvCertificados.Columns[2].Width = 300;
            dgvCertificados.Columns[3].Width = 60;
            dgvCertificados.Columns[4].Width = 60;
            dgvCertificados.Columns[6].Width = 100;
            ColorirGrid();
        }
        private void DadosDataTable(IEnumerable<Certificados> certificados)
        {

            dtTable.Clear();
            foreach (Certificados certificado in certificados)
            {
                dtTable.Rows.Add(certificado.Empresa.First().Codigo, certificado.Empresa.First().Filial, certificado.Empresa.First().NomeEmpresa,
                                    certificado.Valido, certificado.Tipo, certificado.Certificadora, certificado.DataDeVencimento);
            }
        }
        private void ColorirGrid()
        {
            foreach (DataGridViewRow row in dgvCertificados.Rows)
            {
                if (row.Cells[3].Value.ToString().Equals("Ativo"))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.IndianRed;
                }
            }
        }

        private void OrdenarGrid(object sender, DataGridViewCellMouseEventArgs e) => ColorirGrid();

        private void EmissaoCSV(object sender, EventArgs e)
        {
            if (!PossuiItemNaGrid())
                return;

            SaveFileDialog salvar = new SaveFileDialog();
            salvar.FileName = "Relatorio_Parcelamentos_Empresa";

            Microsoft.Office.Interop.Excel.Application App;
            Workbook WorkBook;
            Worksheet WorkSheet;
            object misValue = System.Reflection.Missing.Value;

            App = new Microsoft.Office.Interop.Excel.Application();
            WorkBook = App.Workbooks.Add(misValue);
            WorkSheet = (Worksheet)WorkBook.Worksheets.get_Item(1);
            int linha = 1;
            int coluna = 0;

            DataGridViewCell cell = dgvCertificados[coluna, linha];
            WorkSheet.Cells[linha, 1] = "COD EMPRESA";
            WorkSheet.Cells[linha, 2] = "FILIAL";
            WorkSheet.Cells[linha, 3] = "NOME EMPRESA";
            WorkSheet.Cells[linha, 4] = "VALIDO";
            WorkSheet.Cells[linha, 5] = "TIPO";
            WorkSheet.Cells[linha, 6] = "CERTIFICADORA";
            WorkSheet.Cells[linha, 7] = "DATA VENCIMENTO";

            var codigoEmpresa = WorkSheet.Cells[linha, 1];
            var filialEmpresa = WorkSheet.Cells[linha, 2];
            var nomeEmpresa = WorkSheet.Cells[linha, 3];
            var valido = WorkSheet.Cells[linha, 4];
            var tipo = WorkSheet.Cells[linha, 5];
            var certificadora = WorkSheet.Cells[linha, 6];
            var dataVencimento = WorkSheet.Cells[linha, 7];

            codigoEmpresa.Cells.Font.Bold = true;
            filialEmpresa.Cells.Font.Bold = true;
            nomeEmpresa.Cells.Font.Bold = true;
            valido.Cells.Font.Bold = true;
            tipo.Cells.Font.Bold = true;
            certificadora.Cells.Font.Bold = true;
            dataVencimento.Cells.Font.Bold = true;

            for (linha = 1; linha <= dgvCertificados.RowCount - 1; linha++)
            {
                for (coluna = 0; coluna <= dgvCertificados.ColumnCount - 1; coluna++)
                {
                    DataGridViewCell cell2 = dgvCertificados[coluna, linha - 1];
                    WorkSheet.Cells[linha + 1, coluna + 1] = cell2.Value;
                    var tes = WorkSheet.Cells[linha + 1, coluna + 1];
                    tes.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                }
            }

            Range celulas, celulasTitulo;

            celulasTitulo = WorkSheet.get_Range("A1:K1");
            celulas = WorkSheet.get_Range("A1:Z1000");

            celulasTitulo.Font.Bold = true;
            celulasTitulo.Font.Color = ColorTranslator.ToWin32(Color.Black);
            //celulasTitulo.Interior.Color = ColorTranslator.ToWin32(Color.DarkGray);
            celulas.HorizontalAlignment = Constants.xlCenter;
            celulas.EntireColumn.AutoFit();

            salvar.Title = "Exportar para Excel";
            salvar.Filter = "Arquivo do Excel *.xls | *.xls";
            salvar.ShowDialog();

            WorkBook.SaveAs(salvar.FileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            WorkBook.Close(true, misValue, misValue);
            App.Quit();
        }
        public bool PossuiItemNaGrid()
        {
            if (dgvCertificados.Rows.Count < 1)
            {
                MessageBox.Show("Nenhum dado para ser emitido!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private void EmissaoPDF(object sender, EventArgs e)
        {
            if (!PossuiItemNaGrid())
            {
                return;
            }

            if (repositorioCertificados.QuantidadeCertificados > 18)
                TotalPaginas += (int)Math.Ceiling(
                    (repositorioCertificados.QuantidadeCertificados - 18) / 20F);

            List<Certificados> listaCertificados = repositorioCertificados.Certificados;

            BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            Document documento = new Document(PageSize.A4);
            documento.SetMargins(40, 40, 80, 40);
            documento.AddCreationDate();

            string caminhoArquivo = $@"{Path.GetTempFileName()}" + "relatorio.pdf";

            PdfWriter writer = PdfWriter.GetInstance(documento, new FileStream(caminhoArquivo, FileMode.Create));
            writer.PageEvent = new EventosPagina(TotalPaginas);
            documento.Open();

            Paragraph titulo = new Paragraph();
            titulo.Font = new Fonte(Fonte.FontFamily.TIMES_ROMAN, 16);
            titulo.Alignment = Element.ALIGN_CENTER;
            titulo.Add("RELATÓRIO DE CERTIFICADOS \n \n \n");

            PdfPTable tabela = new PdfPTable(6);

            float[] larguraColunas = { 0.6f, 3.0f, 0.8f, 0.6f, 1.0f, 1.0f };
            tabela.SetWidths(larguraColunas);
            tabela.SetWidths(larguraColunas);
            tabela.DefaultCell.BorderWidth = 0;
            tabela.WidthPercentage = 100;

            CriarCelulaTexto(tabela, "Cód.", 9, true);
            CriarCelulaTexto(tabela, "Nome Empresa", 9, true);
            CriarCelulaTexto(tabela, "Valido", 9, true);
            CriarCelulaTexto(tabela, "Tipo", 9, true);
            CriarCelulaTexto(tabela, "Certificadora", 9, true);
            CriarCelulaTexto(tabela, "Vencimento", 9, true);

            foreach (Certificados certificado in listaCertificados)
            {
                CriarCelulaTexto(tabela, certificado.Empresa.First().Codigo.ToString() + '-' + certificado.Empresa.First().Filial.ToString(), 8, false);
                CriarCelulaTexto(tabela, certificado.Empresa.First().NomeEmpresa.ToString(), 8, false);
                CriarCelulaTexto(tabela, certificado.Valido.ToString(), 8, false);
                CriarCelulaTexto(tabela, certificado.Tipo, 8, false);
                CriarCelulaTexto(tabela, certificado.Certificadora.ToString(), 8, false);
                CriarCelulaTexto(tabela, certificado.DataDeVencimento.Date.ToString().Substring(0, 10), 8, false);
                
            };

            documento.Add(titulo);
            documento.Add(tabela);
            documento.Close();

            System.Diagnostics.Process.Start(caminhoArquivo);
        }

        private static void CriarCelulaTexto(PdfPTable tabela, string texto, int tamanhoFonte, bool negrito,
           int alinhamentoHoriz = PdfPCell.ALIGN_LEFT, bool italico = false, int alturaCelula = 35)
        {
            {
                int estilo = Fonte.NORMAL;
                if (negrito && italico)
                {
                    estilo = Fonte.BOLDITALIC;
                }
                else if (negrito)
                {
                    estilo = Fonte.BOLD;
                }
                else if (italico)
                {
                    estilo = Fonte.ITALIC;
                }

                BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                Fonte fonte = new Fonte(fonteBase, tamanhoFonte,
                    estilo, BaseColor.BLACK);

                var bgColor = BaseColor.WHITE;
                if (tabela.Rows.Count % 2 == 1)
                    bgColor = new BaseColor(0.95f, 0.95f, 0.95f);

                PdfPCell celula = new PdfPCell(new Phrase(texto, fonte))
                {
                    HorizontalAlignment = alinhamentoHoriz,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Border = 0,
                    BorderWidthBottom = 1,
                    FixedHeight = alturaCelula,
                    BackgroundColor = bgColor
                };
                tabela.AddCell(celula);
            }
        }
    }
}
