using geraçãoRelatorioPDF;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.Diagnostics;
using System.Text.Json;

namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static List<Pessoa> pessoaList = new List<Pessoa>();

        static BaseFont fonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false); // fonte base

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            DesserializarPessoas();

            GerarRelatorioEmPdf(100);
        }

        static void DesserializarPessoas()
        {
            if (File.Exists("pessoas.json"))
            {
                using (var sr = new StreamReader("pessoas.json"))
                {
                    var dados = sr.ReadToEnd();
                    pessoaList = JsonSerializer.Deserialize(dados, typeof(List<Pessoa>)) as List<Pessoa>;
                }
            }
        }

        static void GerarRelatorioEmPdf(int qtdePessoas)
        {
            var pessoasSelecionadas = pessoaList.Take(qtdePessoas).ToList();

            if (pessoasSelecionadas.Count > 0)
            {
                // configuração do documento pdf:
                var pxPorMm = 72 / 25.2F; 
                var pdf = new Document(PageSize.A4, 15 * pxPorMm, 15 * pxPorMm, 15 * pxPorMm, 20 * pxPorMm); // criando o pdf com as larguras nas bordas
                var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")}.pdf"; // gerando o nome do arquivo com data e hr quando for criado.
                var arquivo = new FileStream(nomeArquivo, FileMode.Create);
                var write = PdfWriter.GetInstance(pdf, arquivo); // juntando onde o arquivo vat ser 

                pdf.Open(); // abrindo o arquivo

                //adiçao do titulo
                var fonteParagrafo = new Font(fonteBase, 32, Font.NORMAL, BaseColor.Black); // definindo uma fonte padrao para os paragrafos
                var titulo = new Paragraph("Relatórios de Pessoas\n\n", fonteParagrafo); // criando um paragrafo passando a fonte
                titulo.Alignment = Element.ALIGN_LEFT; // estilizando o paragrafo ao meio do pdf
                titulo.SpacingAfter = 4; // colocando um espaço em baixo do meu titulo
                pdf.Add(titulo); // add o paragrafo ao documento pdf.

                // adição da imagem 
                var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img\\youtube.png");
                if (File.Exists(caminhoImagem))
                {
                    Image logo = Image.GetInstance(caminhoImagem);
                    float razaoAlturaLargura = logo.Width / logo.Height;
                    float alturaLogo = 32;
                    float larguraLogo = alturaLogo * razaoAlturaLargura;
                    logo.ScaleToFit(larguraLogo, alturaLogo);
                    var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                    var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                    logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                    write.DirectContent.AddImage(logo, false);
                }

                // adição/criando a tabela de dados:
                var tabela = new PdfPTable(5); // criando uma tabela com 5 colunas.
                float[] larguraColunasTabela = { 0.6f, 2f, 1.5f, 1f, 1f };
                tabela.SetTotalWidth(larguraColunasTabela); // definindo o tamanho das larguras entre as colunas.
                tabela.DefaultCell.BorderWidth = 0; // defindo que as células não teram bordas.
                tabela.WidthPercentage = 100; //definindo que a tabela usará 100% da largura disponóvel da página.

                // criando 5 celular/nome dos campos
                CriarCelulaTexto(tabela, "Código", PdfPCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Nome", PdfPCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Profissão", PdfPCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Salario", PdfPCell.ALIGN_CENTER, true);
                CriarCelulaTexto(tabela, "Empregado", PdfPCell.ALIGN_CENTER, true);

                foreach (var p in pessoasSelecionadas)
                {
                    CriarCelulaTexto(tabela, p.IdPessoa.ToString("D6"), PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Nome + " " + p.Sobrenome, PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Profissao.Nome, PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Salario.ToString("C2"), PdfPCell.ALIGN_CENTER);
                    CriarCelulaTexto(tabela, p.Empregado ? "Sim" : "Não", PdfPCell.ALIGN_CENTER);
                }

                pdf.Add(tabela);

                pdf.Close();
                arquivo.Close();

                // abre o PDF no visualizador padrão:
                var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo);
                if(File.Exists(caminhoPDF))
                {
                    Process.Start(new ProcessStartInfo()
                    {
                        Arguments = $"/c start {caminhoPDF}",
                        FileName = "cmd.exe",
                        CreateNoWindow = true
                    });
                }
            }
        }

        // criando celulas dinamicas
        static void CriarCelulaTexto(PdfPTable tabela, string texto, int alinhamentoHorz = PdfPCell.ALIGN_LEFT, bool negrito = false, bool italico = false, int tamanhoFonte = 12, int alturaCelula = 25)
        {
            var estilo = Font.NORMAL;
            if (negrito && italico)
            {
                estilo = Font.BOLDITALIC;
            }
            else if (negrito)
            {
                estilo = Font.BOLD;
            }
            else if (italico)
            {
                estilo = Font.ITALIC;
            }

            // add uma linha de cada cor, para vizualizar melhor
            var bgColor = BaseColor.White;
            if (tabela.Rows.Count % 2 == 1)
            {
                bgColor = new BaseColor(0.95F, 0.95F, 0.95F);
            }

            var fonteCelula = new Font(fonteBase, 12, estilo, BaseColor.Black);

            var celula = new PdfPCell(new Phrase(texto, fonteCelula));
            celula.HorizontalAlignment = alinhamentoHorz;
            celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            celula.Border = 0;
            celula.BorderWidthBottom = 1;
            celula.FixedHeight = alturaCelula;
            celula.BackgroundColor = bgColor; // add a cor nas celulas
            celula.PaddingBottom = 5; // add uma margem inferior
            tabela.AddCell(celula); // adicionando a celula dentro da tabela
        }
    }
}