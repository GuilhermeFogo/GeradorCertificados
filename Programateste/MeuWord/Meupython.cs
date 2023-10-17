using System;
using System.Diagnostics;
namespace Programateste.MeuWord
{
    public static class Meupython
    {
        public static void ExecutarScript(string argumento)
        {
            try
            {
                // Caminho para o executável do programa secundário
                string pathToProgram = @"..\Script python\dist\main\main.exe";

                // Iniciar um novo processo
                Process.Start(pathToProgram, argumento);

                var processo = new Process
                {
                    StartInfo =
                    {
                        FileName = pathToProgram,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Arguments = argumento
                    }
                };

                // Inicia o processo e aguarda sua conclusão
                processo.Start();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro ao iniciar o programa secundário: " + ex.Message);
            }
        }
    }
}