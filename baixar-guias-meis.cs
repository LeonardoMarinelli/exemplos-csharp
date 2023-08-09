using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Net;
using WindowsInput;
using WindowsInput.Native;
using PdfSharp.Pdf.IO;
using iTextSharp.text.pdf.parser;
using GetDASMEIs.Helpers;
using GetDASMEIs.Entities;

namespace GetDASMEIs.Controllers
{
    public class Fazenda
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string? className, string windowName);
        private static readonly string[] headerSucess = {"CNPJ", "RAZÃO_SOCIAL", "CAMINHO_DO_ARQUIVO"};
        private static readonly string[] headerError = {"CNPJ", "RAZÃO_SOCIAL", "MOTIVO_ERRO"};
        //Essa string dirs contém o caminho onde está a pasta de todos os clientes.
        private static readonly string [] dirs = Directory.GetDirectories(@"CAMINHO_A_SER_SALVO");
        static private readonly string dateForExcel = DateTime.UtcNow.ToString().Replace("/", "").Replace(" ", "").Replace(":", "");

        //O método recebe 3 strings, sendo o CNPJ sem pontuação, dataMes que seria "Agosto/2023" por exemplo e o mês, que seria "08".
        //Esse método GetDocs pode ser usado em loop, por exemplo, com um array de string que contenha vários CNPJs, porém deve-se deixar ele acabar para clicar/usar o teclado.
        //Como o site da Receita Federal é protegido pelo hCaptcha, não foi possível fazer via requisição com RestSharp e nem com Selenium.
        //No começo do método GetDocs, indicamos exatamente a resolução que a janela seria aberta, impedindo qualquer problema de dimensão.
        public static async Task GetDocs(string cnpj, string dataMes, string mes) 
        {
            //Verifica se o site está online e funcionando.
            while (true) 
            {
                bool ok = await IsWorking();
                if (ok) {
                    break;
                } else {
                    Thread.Sleep(15000);
                }
            }
            Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", 
            "https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao --incognito --window-size=1000,720 --window-position=100,60");

            IntPtr handle = FindWindow(null, "PGMEI - Programa Gerador de DAS do Microempreendedor Individual");
            SwitchToThisWindow(handle, true);

            Thread.Sleep(1300);
            ChromeController(cnpj);
            Thread.Sleep(5000);

            string sourcePdfPath = @$"CAMINHO_QUE_SERÁ_SALVO_O_ARQUIVO + NOME_DO_ARQUIVO";
            PdfSharp.Pdf.PdfDocument sourceDocument = PdfReader.Open(sourcePdfPath, PdfDocumentOpenMode.Import);
            int selectedPageNumber = FindTextInPDF(sourcePdfPath, dataMes);
            PdfSharp.Pdf.PdfDocument newDocument = new();

            //Caso o retorno tenha sido diferente de -1 (que significa não encontrado), ele segue. Se não, ele imprime uma planilha Excel indicando o erro.
            if (selectedPageNumber > 0) 
            {
                if (!Directory.Exists("./data/temp")) Directory.CreateDirectory("./data/temp"); 

                PdfSharp.Pdf.PdfPage selectedPage = sourceDocument.Pages[selectedPageNumber - 1];
                newDocument.AddPage(selectedPage);
                string path = @$"{Directory.GetCurrentDirectory()}\data\temp\DAS-PGMEI - {mes}.{dataMes.Split("/")[1]}.pdf";

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                newDocument.Save(path);
                string clientName = GetNameFromPDF(PdfContent.Get(path));
                File.Copy(path, @$"{Directory.GetCurrentDirectory()}\data\temp\DAS-PGMEI - {clientName} {mes}.{dataMes.Split("/")[1]}.pdf", true);
                File.Delete(path);

                path = @$"{Directory.GetCurrentDirectory()}\data\temp\DAS-PGMEI - {clientName} {mes}.{dataMes.Split("/")[1]}.pdf";
                CopyFiles(path, clientName, $"{mes}-{dataMes.Split("/")[0].ToUpper()}", $"DAS-PGMEI - {clientName} {mes}.{dataMes.Split("/")[1]}.pdf", cnpj, clientName);
            } else {
                ClientsError error = new() 
                {
                    Cnpj = cnpj,
                    Name = GetNameFromPDF(PdfContent.Get(sourcePdfPath)),
                    Reason = $"Não foi possível localizar a guia do mês {dataMes} no PDF"
                };
                //Acessa um método que escreve uma planilha em xlsx, indicando os erros ou sucessos ao emitir. Verifica se já existe o header, se não ele não escreve novamente.
                //Será usado bastante, e está contido nos Helpers.
                Excel.WriteXLSX(error, headerError, @$"ERROS - OBTER GUIA DAS {mes}.{dataMes.Split("/")[1]} - {dateForExcel}");
            }

            //Mata todos os processos do Chrome. O site da receita armazena os dados, logo será impossível executar esse método várias vezes na mesma janela.
            //O ideal é matar os processos do Chrome e abrir novamente, caso o método esteja em loop. Também deleta os lixos que não serão mais usados.
            Process[] chromeProcesses = Process.GetProcessesByName("chrome");
            foreach (Process process in chromeProcesses) process.Kill();
            File.Delete(sourcePdfPath);
        }

        //Método que copia os arquivos para o caminho desejado, caso seja necessário.
        static void CopyFiles(string path, string clientName, string monthName, string fileName, string cnpj, string razaoSocial) 
        {
            bool found = false;
            foreach (string dir in dirs) 
            {
                if (dir.Contains(clientName)) 
                {
                    found = true;
                    if (!Directory.Exists($@"CAMINHO_A_SER_COPIADO_NO_FIM")) Directory.CreateDirectory($@"CAMINHO_A_SER_COPIADO_NO_FIM"); 
                   
                    string pathToCopy = $@"CAMINHO_A_SER_COPIADO_NO_FIM + NOME_DO_ARQUIVO";
                    File.Copy(path, pathToCopy, true);
                    File.Delete(path);
                    
                    ClientsSucess success = new () 
                    {
                        Cnpj = cnpj,
                        Name = razaoSocial,
                        Path = pathToCopy
                    };
                    Excel.WriteXLSX(success, headerSucess,  @$"SUCESSO - PROCESSO GUIA DAS {monthName.Replace("-", ".")} - {dateForExcel}");
                    continue;
                }
            }
            if (!found) 
            {
                ClientsError error = new () 
                {
                    Cnpj = cnpj,
                    Name = razaoSocial,
                    Reason = "Não foi possível localizar a pasta do cliente na ISO"
                };
                Excel.WriteXLSX(error, headerError, @$"ERROS - COPIAR GUIA DAS {monthName.Replace("-", ".")} - {dateForExcel}");
            }
        }

        //Método que verifica se o site está funcionando adequadamente.
        static async Task<bool> IsWorking() 
        {
            string url = "https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao";
            using HttpClient httpClient = new();
            HttpResponseMessage response = await httpClient.GetAsync(url);
            return response.StatusCode == HttpStatusCode.OK;
        }

        //O método retorna a Razão Social contida na guia achada utilizando RegEx.
        static string GetNameFromPDF(string textPDF) 
        {
            string pattern = @"\b[A-Z\s]+\d{11}\b";
            MatchCollection matches = Regex.Matches(textPDF, pattern);
            return matches[0].Value.Trim();
        }

        //Esse método serve para localizar qual das páginas baixadas será a utilizada (guia do mês desejado). Se ele não achar, retorna -1.
        static int FindTextInPDF(string filePath, string searchText) 
        {
            using iTextSharp.text.pdf.PdfReader reader = new(filePath);
            for (int page = 1; page <= reader.NumberOfPages; page++) 
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                string currentPageText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);

                if (currentPageText.Contains(searchText)) return page;
            }
            return -1;
        }

        //Método que controla o mouse e teclado no site da Receita. 
        static void ChromeController(string cnpj) {
            InputSimulator simulator = new();
            Thread.Sleep(800);
            simulator.Keyboard.TextEntry(cnpj);
            Thread.Sleep(800);
            simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(4000);
            simulator.Mouse.MoveMouseTo(11500, 20000);
            simulator.Mouse.LeftButtonClick();
            Thread.Sleep(800);
            simulator.Mouse.MoveMouseTo(29000, 29500);
            simulator.Mouse.LeftButtonClick();
            simulator.Mouse.MoveMouseTo(29000, 44600);
            simulator.Mouse.LeftButtonClick();
            simulator.Mouse.MoveMouseTo(31500, 29500);
            simulator.Mouse.LeftButtonClick();
            Thread.Sleep(800);
            simulator.Mouse.MoveMouseTo(9000, 38000);
            simulator.Mouse.LeftButtonClick();
            simulator.Mouse.VerticalScroll(-10);
            Thread.Sleep(600);
            simulator.Mouse.MoveMouseTo(27000, 33000);
            simulator.Mouse.LeftButtonClick();
            Thread.Sleep(5500);
            simulator.Mouse.VerticalScroll(-10);
            Thread.Sleep(600);
            simulator.Mouse.MoveMouseTo(24500, 38500);
            simulator.Mouse.LeftButtonClick();
            Thread.Sleep(1600);
            simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }
    }
}