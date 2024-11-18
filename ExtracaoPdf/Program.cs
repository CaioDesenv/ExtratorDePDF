using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.IO; // Para manipular arquivos e pastas
using System.Drawing; // Para tipos gráficos como Size, Point, Font
using System.Windows.Forms; // Para a interface gráfica
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using ClosedXML.Excel;

namespace ExtracaoPdf
{
    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private Button btnSelectFolder;
        private Button btnStartProcess;
        private Button btnExit;
        private Label lblProgress;
        private FolderBrowserDialog folderDialog;
        private string selectedFolderPath;

        public MainForm()
        {
            // Configuração básica da janela
            this.Text = "Processador de PDFs";
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Label "Progresso"
            lblProgress = new Label
            {
                Text = "Progresso: Aguardando ação do usuário...",
                AutoSize = true,
                Location = new Point(20, 20),
                Font = new Font("Arial", 10, FontStyle.Bold)
            };
            this.Controls.Add(lblProgress);

            // Botão "Selecionar Pasta"
            btnSelectFolder = new Button
            {
                Text = "Selecionar Pasta",
                Location = new Point(20, 60),
                Size = new Size(150, 30)
            };
            btnSelectFolder.Click += BtnSelectFolder_Click;
            this.Controls.Add(btnSelectFolder);

            // Botão "Iniciar Processo"
            btnStartProcess = new Button
            {
                Text = "Iniciar Processo",
                Location = new Point(20, 110),
                Size = new Size(150, 30),
                Enabled = false // Ativado somente após a seleção de uma pasta
            };
            btnStartProcess.Click += BtnStartProcess_Click;
            this.Controls.Add(btnStartProcess);

            // Botão "Sair"
            btnExit = new Button
            {
                Text = "Sair",
                Location = new Point(20, 160),
                Size = new Size(150, 30)
            };
            btnExit.Click += (sender, e) => Application.Exit();
            this.Controls.Add(btnExit);

            // Diálogo para selecionar pasta
            folderDialog = new FolderBrowserDialog();
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFolderPath = folderDialog.SelectedPath;
                lblProgress.Text = $"Pasta selecionada: {selectedFolderPath}";
                btnStartProcess.Enabled = true; // Habilita o botão "Iniciar Processo"
            }
        }

        private void BtnStartProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("Por favor, selecione uma pasta antes de iniciar o processo.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            lblProgress.Text = "Progresso: Processando PDFs...";
            Application.DoEvents(); // Atualiza a interface

            string excelPath = System.IO.Path.Combine(selectedFolderPath, "Resultado-Extracao.xlsx");

            using (var workbook = new XLWorkbook())
            {
                foreach (var pdfFile in Directory.GetFiles(selectedFolderPath, "*.pdf"))
                {
                    try
                    {
                        lblProgress.Text = $"Processando: {System.IO.Path.GetFileName(pdfFile)}...";
                        Application.DoEvents();

                        string conteudoCompleto = ExtrairTextoCompletoDoPdf(pdfFile);
                        string informacoesEssenciaisTexto = ExtrairInformacoesEssenciais(conteudoCompleto);
                        string[] dadosBrutos = informacoesEssenciaisTexto.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        List<string> valoresPosicao2 = new List<string>(dadosBrutos[2].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
                        InformacoesEssenciais informacoesEssenciais = ProcessarInformacoesEssenciais(informacoesEssenciaisTexto);
                        string informacoesFinanceirasBrutas = ExtrairInformacoesFinanceirasBrutas(informacoesEssenciaisTexto);
                        string[] valoresFinanceiros = ExtrairValoresParaVetor(informacoesFinanceirasBrutas);
                        List<string[]> matrizTitulos = ExtrairMatrizTitulos(conteudoCompleto);
                        decimal resultadoDeducao = CalcularDeducaoTotal(valoresFinanceiros);

                        string sheetName = System.IO.Path.GetFileNameWithoutExtension(pdfFile);
                        ExportarParaExcel(workbook, matrizTitulos, informacoesEssenciais, valoresFinanceiros, resultadoDeducao, valoresPosicao2, sheetName);
                    } catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao processar o arquivo {System.IO.Path.GetFileName(pdfFile)}: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                lblProgress.Text = "Progresso: Exportando para Excel...";
                Application.DoEvents();

                workbook.SaveAs(excelPath);

                lblProgress.Text = "Progresso: Finalizado.";
                MessageBox.Show($"Arquivo Excel gerado com sucesso em: {excelPath}", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        static void ExportarParaExcel(XLWorkbook workbook, List<string[]> matrizTitulos, InformacoesEssenciais informacoesEssenciais, string[] valoresFinanceiros, decimal resultadoDeducao, List<string> valoresPosicao2, string sheetName)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Adiciona o cabeçalho da Tabela de Títulos
            worksheet.Cell(1, 1).Value = "Tabela de Títulos";
            worksheet.Cell(2, 1).Value = "Seq.";
            worksheet.Cell(2, 2).Value = "Sacado";
            worksheet.Cell(2, 3).Value = "CPF/CNPJ";
            worksheet.Cell(2, 4).Value = "S. Número";
            worksheet.Cell(2, 5).Value = "Emissão";
            worksheet.Cell(2, 6).Value = "Tipo";
            worksheet.Cell(2, 7).Value = "Aceite";
            worksheet.Cell(2, 8).Value = "Valor Título";
            worksheet.Cell(2, 9).Value = "Vencimento";

            // Insere os dados da Tabela de Títulos
            int currentRow = 3;
            foreach (var linha in matrizTitulos)
            {
                for (int col = 0; col < linha.Length; col++)
                {
                    worksheet.Cell(currentRow, col + 1).Value = linha[col];
                }
                currentRow++;
            }

            // Adiciona as Informações Essenciais
            int infoStartRow = matrizTitulos.Count + 5;
            worksheet.Cell(infoStartRow, 1).Value = "Informações Essenciais Extraídas";
            worksheet.Cell(infoStartRow + 1, 1).Value = "Data da Operação";
            worksheet.Cell(infoStartRow + 1, 2).Value = informacoesEssenciais.DataOperacao;
            worksheet.Cell(infoStartRow + 2, 1).Value = "CPF/CNPJ";
            worksheet.Cell(infoStartRow + 2, 2).Value = informacoesEssenciais.CPFCNPJ;

            worksheet.Cell(infoStartRow + 3, 1).Value = "Cliente/Cedente";
            string clienteCedente = string.Join(" ", valoresPosicao2.GetRange(4, valoresPosicao2.Count - 4));
            worksheet.Cell(infoStartRow + 3, 2).Value = clienteCedente;

            worksheet.Cell(infoStartRow + 4, 1).Value = "Valor Total do(s) Título(s)";
            worksheet.Cell(infoStartRow + 4, 2).Value = valoresFinanceiros[0];
            worksheet.Cell(infoStartRow + 5, 1).Value = "Valor Líquido";
            worksheet.Cell(infoStartRow + 5, 2).Value = valoresFinanceiros[5];
            worksheet.Cell(infoStartRow + 6, 1).Value = "Taxa";
            worksheet.Cell(infoStartRow + 6, 2).Value = valoresFinanceiros[6];
            worksheet.Cell(infoStartRow + 7, 1).Value = "Custo Efetivo Total (CET)";
            worksheet.Cell(infoStartRow + 7, 2).Value = $"{valoresFinanceiros[9]} / {valoresFinanceiros[12]}";
            worksheet.Cell(infoStartRow + 8, 1).Value = "Valor da Tarifa";
            worksheet.Cell(infoStartRow + 8, 2).Value = valoresFinanceiros[15];
            worksheet.Cell(infoStartRow + 9, 1).Value = "Valor IOF";
            worksheet.Cell(infoStartRow + 9, 2).Value = valoresFinanceiros[16];

            // Adiciona o Resultado da Dedução
            int resultadoStartRow = infoStartRow + 11;
            worksheet.Cell(resultadoStartRow, 1).Value = "Resultado da Dedução";
            worksheet.Cell(resultadoStartRow + 1, 1).Value = "Valor Total do(s) Título(s)";
            worksheet.Cell(resultadoStartRow + 1, 2).Value = valoresFinanceiros[0];
            worksheet.Cell(resultadoStartRow + 2, 1).Value = "Deduzindo Valor Líquido";
            worksheet.Cell(resultadoStartRow + 2, 2).Value = valoresFinanceiros[5];
            worksheet.Cell(resultadoStartRow + 3, 1).Value = "Deduzindo Valor da Tarifa";
            worksheet.Cell(resultadoStartRow + 3, 2).Value = valoresFinanceiros[15];
            worksheet.Cell(resultadoStartRow + 4, 1).Value = "Deduzindo Valor IOF";
            worksheet.Cell(resultadoStartRow + 4, 2).Value = valoresFinanceiros[16];
            worksheet.Cell(resultadoStartRow + 5, 1).Value = "Resultado";
            worksheet.Cell(resultadoStartRow + 5, 2).Value = resultadoDeducao.ToString("C", new CultureInfo("pt-BR"));

            // Ajusta a largura das colunas
            worksheet.Columns().AdjustToContents();
        }


        static decimal CalcularDeducaoTotal(string[] valoresFinanceiros)
        {
            var culture = new CultureInfo("pt-BR");
            decimal valorTotal = decimal.Parse(valoresFinanceiros[0], culture);
            decimal valorLiquido = decimal.Parse(valoresFinanceiros[5], culture);
            decimal valorTarifa = decimal.Parse(valoresFinanceiros[15], culture);
            decimal valorIOF = decimal.Parse(valoresFinanceiros[16], culture);
            return valorTotal - (valorLiquido + valorTarifa + valorIOF);
        }

        static string ExtrairTextoCompletoDoPdf(string caminhoPdf)
        {
            StringBuilder conteudoCompleto = new StringBuilder();
            var reader = new PdfReader(caminhoPdf);
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                string textoPagina = PdfTextExtractor.GetTextFromPage(reader, i);
                textoPagina = RemoverTextoIndesejado(textoPagina);
                conteudoCompleto.AppendLine(textoPagina);
            }
            return conteudoCompleto.ToString();
        }

        static string RemoverTextoIndesejado(string texto)
        {
            string pattern = @"Atendimento personalizado.*Demais localidades|CENTRAL DE SUPORTE.*OUVIDORIA|Entregamos nesta data.*Custo Efetivo Total abaixo indicados\.";
            return Regex.Replace(texto, pattern, "", RegexOptions.Singleline);
        }

        static string ExtrairInformacoesEssenciais(string conteudoCompleto)
        {
            string pattern = @"Data da Operação Canal\s*(.*?)\s*RELAÇÃO DO\(S\) TÍTULO\(S\) PARA DESCONTO";
            var match = Regex.Match(conteudoCompleto, pattern, RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value.Trim() : "Informações essenciais não encontradas.";
        }

        static List<string[]> ExtrairMatrizTitulos(string conteudoCompleto)
        {
            string pattern = @"(\d+)\s+([\w\s]+)\s+([\d./-]+)\s+([\d\w-]+)\s+([\d/]+)\s+(DM|NP|CH)\s+(Sim|Não)\s+([\d.,]+)\s+([\d/]+)";
            var matches = Regex.Matches(conteudoCompleto, pattern);
            List<string[]> matrizTitulos = new List<string[]>();
            foreach (Match match in matches)
            {
                matrizTitulos.Add(new[]
                {
                    match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value, match.Groups[4].Value,
                    match.Groups[5].Value, match.Groups[6].Value, match.Groups[7].Value, match.Groups[8].Value,
                    match.Groups[9].Value
                });
            }
            return matrizTitulos;
        }

        static InformacoesEssenciais ProcessarInformacoesEssenciais(string informacoesEssenciaisTexto)
        {
            var informacoes = new InformacoesEssenciais
            {
                DataOperacao = Regex.Match(informacoesEssenciaisTexto, @"\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}").Value,
                AgenciaContaCredito = Regex.Match(informacoesEssenciaisTexto, @"\d{4} / \d{12}-\d").Value,
                CPFCNPJ = Regex.Match(informacoesEssenciaisTexto, @"\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}").Value,
                ClienteCedente = Regex.Match(informacoesEssenciaisTexto, @"Cliente/Cedente\s+([\w\s]+)")?.Groups[1]?.Value.Trim() ?? "Não encontrado",
                ValorTotal = Regex.Match(informacoesEssenciaisTexto, @"Valor Total.*?\s+([\d.,]+)").Groups[1].Value,
                ValorLiquido = Regex.Match(informacoesEssenciaisTexto, @"Valor Líquido - R\$\s+([\d.,]+)").Groups[1].Value,
                Taxa = Regex.Match(informacoesEssenciaisTexto, @"Taxa\s+([\d.,%]+)").Groups[1].Value,
                CustoEfetivoTotal = Regex.Match(informacoesEssenciaisTexto, @"CET\s+([\d.,%]+)").Groups[1].Value,
                ValorTarifa = Regex.Match(informacoesEssenciaisTexto, @"Valor da Tarifa - R\$\s+([\d.,]+)").Groups[1].Value,
                ValorIOF = Regex.Match(informacoesEssenciaisTexto, @"Valor IOF - R\$\s+([\d.,]+)").Groups[1].Value
            };
            return informacoes;
        }

        static string ExtrairInformacoesFinanceirasBrutas(string informacoesEssenciaisTexto)
        {
            string pattern = @"Valor Total do\(s\) Título\(s\) R\$ Qtde Título\(s\) Vencimento Final\s*(.*)";
            var match = Regex.Match(informacoesEssenciaisTexto, pattern, RegexOptions.Singleline);
            if (match.Success)
            {
                string informacoesFinanceirasBrutas = match.Groups[1].Value.Trim();
                string lixoPattern = @"Valor Líquido - R\$ Taxa Custo Efetivo Total \(CET\) Valor da Tarifa - R\$ Valor IOF - R\$";
                return Regex.Replace(informacoesFinanceirasBrutas, lixoPattern, "", RegexOptions.Singleline).Trim();
            }
            return "Informações financeiras brutas não encontradas.";
        }

        static string[] ExtrairValoresParaVetor(string informacoesFinanceirasBrutas)
        {
            string pattern = @"[\d.,]+(?:%|\(a\.m\.\)|\(a\.a\.\))?";
            var matches = Regex.Matches(informacoesFinanceirasBrutas, pattern);
            var valores = new List<string>();
            foreach (Match match in matches)
            {
                valores.Add(match.Value);
            }
            return valores.ToArray();
        }
    }

    class InformacoesEssenciais
    {
        public string DataOperacao { get; set; }
        public string AgenciaContaCredito { get; set; }
        public string CPFCNPJ { get; set; }
        public string ClienteCedente { get; set; }
        public string ValorTotal { get; set; }
        public string ValorLiquido { get; set; }
        public string Taxa { get; set; }
        public string CustoEfetivoTotal { get; set; }
        public string ValorTarifa { get; set; }
        public string ValorIOF { get; set; }
    }
}
