using Certificados_Clientes_WL.Formularios;
using Certificados_Clientes_WL.Repositorios;
using Certificados_Clientes_WL.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Certificados_Clientes_WL
{
    public partial class FrmLogin : Form
    {
        private RepositorioUsuario repositorioUsuario = new RepositorioUsuario();
        private Criptografia criptografia = new Criptografia();
        public FrmLogin()
        {
            InitializeComponent();
            MapeiaNomeUsuarios();
        }

        private void BtnClose_Click(object sender, EventArgs e) => Application.Exit();

        private void ImgOlhosFechados_Click(object sender, EventArgs e) => VisualizarSenha();

        private void ImgOlhosAberto_Click(object sender, EventArgs e) => NaoVisualizarSenha();
        private void MapeiaNomeUsuarios()
        {
            foreach (String nomeUsuarios in repositorioUsuario.ListaNomeUsuariosWL())
            {
                comboBoxUsuarios.Items.Add(nomeUsuarios);
            }
            comboBoxUsuarios.SelectedIndex = 0;
        }

        private void Login(object sender, EventArgs e)
        {
            string senha = criptografia.GerarHashMd5(txtSenha.Text);
            if (repositorioUsuario.Logar(comboBoxUsuarios.Text, senha, txtSenha.Text))
            {
                FrmCertificados frmCertificados = new FrmCertificados();
                this.Hide();
                frmCertificados.Show();
            }
            else
            {
                MessageBox.Show("Usuário ou senha incorreto!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
        }

        private void NaoVisualizarSenha()
        {
            txtSenha.UseSystemPasswordChar = true;
            imgOlhosAberto.Visible = false;
            imgOlhosFechados.Visible = true;
        }

        private void VisualizarSenha()
        {
            txtSenha.UseSystemPasswordChar = false;
            imgOlhosAberto.Visible = true;
            imgOlhosFechados.Visible = false;
        }

        private void Timer1_Tick(object sender, EventArgs e) => label2.Text = DateTime.Now.ToString("HH:mm:ss");
    }
}
