using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Programateste.Modal
{
    public class Certificado
    {
        public string Nome { get; set; }
        public string Email { get; set; }
        public string Curso { get; set; }
        public DateTime Data { get; set; }
        public string Palestrante { get; set; }
        public string CargaHoraria { get; set; }

        public override string ToString()
        {
            return  string.Format("{0}, {1}, {2}, {3}, {4}, {5}", Nome,Email,Curso,Data,Palestrante,CargaHoraria);
        }
    }
}
