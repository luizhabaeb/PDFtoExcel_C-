using System;
using System.IO;
using System.Windows.Forms;
using PdfiumViewer;

namespace PDFtoExcel
{
    public partial class Form1 : Form
    {                
        //Instanciação do pacote - PDF Viewer - que possiblita a leitura do documento.
        PdfiumViewer.PdfViewer pdf;

        // Componentes do Programa (design)
        public Form1()
        {
            InitializeComponent();
            pdf = new PdfViewer();
            pdf.Width = this.Width - 10;
            pdf.Height = this.Height - 20;
            this.Controls.Add(pdf);            
        }

        // Método para abrir o File Explorer e selecionar o arquivo a ser convertido
        public void openfile(string filepath)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(filepath);
            var stream = new System.IO.MemoryStream(bytes);
            PdfDocument pdfDocument = PdfDocument.Load(stream);
            pdf.Document = pdfDocument;
        }

        //função completa: Abrir PDF > Exibir no programa > Converter em EXCEL > Abrir e Salvar o arquivo.
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                openfile(dialog.FileName);
            }                       
            string pathToPdf = dialog.FileName;
            string pathToExcel = Path.ChangeExtension(pathToPdf, ".xls");
            

            // Instânciação do pacote para converter as tabelas de PDF em XLS.
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();

            // 'true' = o calendário usado, a ordem de classificação das strings e a formatação de datas e números.
            // 'false' = Ignora os textos e converte apenas os dados que estão tabelados.
            f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = true;

            // 'true'  = Preserva o layout (template original do PDF)
            // 'false' = Define 'tabelas' antes do texto (grades)
            f.ExcelOptions.PreservePageLayout = true;

            // Aqui definimos as informações do sistema de escrita do país (formatação)
            // Calendário usado; a ordem de classificação das strings e a formatação de datas e números.
            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("pt-BR");
            ci.NumberFormat.NumberDecimalSeparator = ",";
            ci.NumberFormat.NumberGroupSeparator = ".";
            f.ExcelOptions.CultureInfo = ci;

            f.OpenPdf(pathToPdf); 

            if (f.PageCount > 0)
            {
                int result = f.ToExcel(pathToExcel);

                // Abre o arquivo gerado em Excel.
                if (result == 0)
                {                
                    System.Diagnostics.Process.Start(pathToExcel);           

                }
            }

        }
    }
}
