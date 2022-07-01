using ClosedXML.Excel;
using Programateste.Modal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Programateste.MeuWord;
using SystemAPI.Mensagero;

namespace Programateste
{
    public class Program
    {
        private static XLWorkbook xls { get; set; }
        public static string diretoriobase { get; set; }
        public static void Main(string[] args)
        {

            bool continua = true;
            while (continua)
            {
                Console.WriteLine("0- Sair do programa ||1 - Executar");
                string escolha = Console.ReadLine();
                switch (escolha)
                {
                    case "0":
                        continua = false;
                        Console.WriteLine("Saindo...");
                        break;
                    case "1":
                        Executar();
                        break;
                    default:
                        Console.WriteLine("Tente Novamente");
                        break;
                }

            }

        }

        private static void Executar()
        {
            Console.WriteLine("Digite seu E-mail");
            string email = Console.ReadLine();
            Console.WriteLine("Senha do Email");
            string senha = Console.ReadLine();

            if (email.Contains("@"))
            {

                Console.WriteLine("Infome o caminho do arquivo Excel base");
                diretoriobase = Console.ReadLine();
                Console.WriteLine("================================================");
                Console.WriteLine("Executando");
                Console.WriteLine("================================================");
                var certificados = TransformaExcelCertificado();
                Console.WriteLine("================================================");
                Console.WriteLine("Arquivo Excel Lido");
                Console.WriteLine("================================================");
                Console.WriteLine("Gerando arquivos Word");
                Console.WriteLine("================================================");

                certificados.ForEach(x =>
                {
                    Words words = new Words($"{x.Nome}", diretoriobase);
                    words.CriandoCertificado(x.ToString());
                });

                Console.WriteLine("================================================");
                Console.WriteLine("Transformando Word em PDF");
                Console.WriteLine("================================================");

                try
                {
                    Directory.Delete($"{diretoriobase}/pdf", true);
                    Directory.CreateDirectory($"{diretoriobase}/pdf");
                }
                catch (Exception e)
                {
                    Directory.CreateDirectory($"{diretoriobase}/pdf");
                }
                IMensageiro mensageiro = new Mensageiro();
                certificados.ForEach(x =>
                {
                    Words words = new Words($"{x.Nome}", diretoriobase);
                    words.WordToPDF();
                    //string assunto = $"[{x.Curso}] Certitificado de Conclusão Mazars";
                    //string mensagem = $"Olá {x.Nome}, aqui segue o certificado do curso: {x.Curso}";
                    //Console.WriteLine($"Enviando e-mail para:{x.Email}");
                    //mensageiro.EnviarEmailHTML(x.Email, assunto, mensagem, diretoriobase + $"{x.Nome}.pdf");
                });
            }
            else
            {
                Console.WriteLine("Digite o email valido");
            }
        }

        private static List<Certificado> TransformaExcelCertificado()
        {
            var certificados = new List<Certificado>();
            try
            {
                xls = new XLWorkbook(diretoriobase + "MALA DIRETA1.xlsx");
                var planilha = xls.Worksheets.First(w => w.Name == "Publico");
                var totalLinhas = planilha.Rows().Count();
                // primeira linha é o cabecalho
                for (int l = 2; l < totalLinhas; l++)
                {
                    if(planilha.Cell($"A{l}").Value !=null || !planilha.Cell($"A{l}").Value.Equals(""))
                    {
                        var certificado = new Certificado
                        {
                            Nome = planilha.Cell($"A{l}").Value.ToString(),
                            Curso = planilha.Cell($"B{l}").Value.ToString(),
                            Data = DateTime.Parse(planilha.Cell($"C{l}").Value.ToString()),
                            CargaHoraria = planilha.Cell($"D{l}").Value.ToString(),
                            Palestrante = planilha.Cell($"E{l}").Value.ToString(),
                            Email = planilha.Cell($"F{l}").Value.ToString()
                        };
                        certificados.Add(certificado);
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("========================================================");
                Console.WriteLine("Digite Novamente o Caminho do excel base");
                string caminho = Console.ReadLine();
                diretoriobase = caminho;
                var cerificados2 = TransformaExcelCertificado();
                return cerificados2;
            }
            return certificados;
        }
    }
}
