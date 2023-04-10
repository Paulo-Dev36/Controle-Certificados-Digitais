using System;
using System.Collections;
using System.Collections.Generic;

namespace Certificados_Clientes_WL.Classes
{
    public class Certificados : IEntidade
    {
        public int Id { get; set; }
        public string NomeCertificado { get; set; }
        public string Valido { get; set; }
        public DateTime DataDeVencimento { get; set; }

        public int CodigoEmpresa { get; set; }

        public int Filial { get; set; }

        public Byte[] Arquivo { get; set; }

        public DateTime DataCadastro { get; set; }
        public IEnumerable<Usuarios> Usuario { get; set; }
        public string Senha { get; set; }
        public IEnumerable<Empresas> Empresa { get; set; }

        public string Tipo { get; set; }
        public string Certificadora { get; set; }

    }
}
