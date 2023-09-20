using DocumentFormat.OpenXml.Packaging;

namespace Programateste.MeuWord
{
    public class Words
    {
        public string NomeArquivo { get; private set; }
        public string Filepath2 { get; private set; }

        public Words()
        {

        }
        public Words(string Nomearquivo, string arquivobase)
        {
            this.NomeArquivo = Nomearquivo;
            this.Filepath2 = arquivobase;
        }

        public void CriandoCertificado(string conteudo)
        {

            // MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            using (WordprocessingDocument certificadoBase = WordprocessingDocument.Open(this.Filepath2 + $"CERTIFICADOS_base.docx", true))
            {
                var array = conteudo.Split(",");
                string textodoc = null;

                textodoc = certificadoBase.MainDocumentPart.Document.InnerXml;

                var textocertificado = textodoc.Replace("campoNome", array[0]).Replace("campoCurso", array[2])
                    .Replace("campoData", array[3]).Replace("campoCarga", array[5]).Replace("campoPalestrante", array[4]);

                var textopadrao = textocertificado.Replace(array[0], "campoNome").Replace(array[2], "campoCurso")
                   .Replace(array[3], "campoData").Replace(array[5], "campoCarga").Replace(array[4], "campoPalestrante");

                certificadoBase.MainDocumentPart.Document.InnerXml = textocertificado;
                certificadoBase.Clone(this.Filepath2 + $"word/{this.NomeArquivo}.docx");
                certificadoBase.MainDocumentPart.Document.InnerXml = textopadrao;
                certificadoBase.Save();
                certificadoBase.Close();


            }
        }

        public void WordToPDF()
        {
            string caminho = this.Filepath2 + $"{this.NomeArquivo}.docx";
            string caminhoPDF = this.Filepath2 + $"/pdf/{this.NomeArquivo}.pdf";

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(caminho);
            //Converter para PDF
            doc.ExportAsFixedFormat(caminhoPDF, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            doc.Close();
            app.Quit();

        }


        //private byte[] ConvertToPDF(HttpPostedFileBase file, string ext)
        //{
        //    //salva o arquivo na pasta App_Data
        //    string path = Server.MapPath($"~/App_Data/nome_arquivo");
        //    file.SaveAs($"{path}{ext}");

        //    //Micrososft Word
        //    if (ext == ".doc" || ext == ".docx")
        //    {
        //        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        //        Microsoft.Office.Interop.Word.Document doc = app.Documents.Open($"{path}{ext}");
        //        //Converter para PDF
        //        doc.ExportAsFixedFormat($"{path}.pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
        //        doc.Close();
        //        app.Quit();
        //        //Leia o arquivo e retorna bytes[]
        //        return System.IO.File.ReadAllBytes($"{path}.pdf");
        //    }
        //    //Microsoft Excel
        //    if (ext == ".xls" || ext == ".xlsx")
        //    {
        //        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open($"{path}{ext}");
        //        //Converter para PDF
        //        wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, $"{path}.pdf");
        //        wkb.Close();
        //        app.Quit();
        //        //Leia o arquivo e retorna bytes[]
        //        return System.IO.File.ReadAllBytes($"{path}.pdf");
        //    }
        //    //Microsoft PowerPoint
        //    else
        //    {
        //        //ppt || pptx 
        //        Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();
        //        Microsoft.Office.Interop.PowerPoint.Presentation presentation = app.Presentations.Open($"{path}{ext}",
        //            Microsoft.Office.Core.MsoTriState.msoTrue,
        //            Microsoft.Office.Core.MsoTriState.msoFalse,
        //            Microsoft.Office.Core.MsoTriState.msoFalse);

        //        //Converter para PDF
        //        presentation.ExportAsFixedFormat($"{path}.pdf", Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
        //        presentation.Close();
        //        app.Quit();
        //        //Leia o arquivo e retorna bytes[]
        //        return System.IO.File.ReadAllBytes($"{path}.pdf");
        //    }
        //}
    }
}
