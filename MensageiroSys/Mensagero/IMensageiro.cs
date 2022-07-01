using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SystemAPI.Mensagero
{
    public interface IMensageiro
    {
        void EnviarEmail(string para, string asssunto, string mensagem);
        void EnviarEmailComAnexo(string para, string asssunto, string mensagem, string nomeArquivo);

        void EnviarEmailHTML(string para, string assunto, string HTML);
        void EnviarEmailHTMLComAnexo(string para, string assunto, string HTML, string nomeArquivo);
    }
}
