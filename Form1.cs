
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;


namespace CreciSP
{
    public partial class Form1 : Form
    {


        List<string> numeroRegistro = new List<string>();

        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        private string _applicationName = "Lista de Corretores - CRECISP";

        private string _spreadsheetId = "1KUWSn5Mr8AzqvL3-TBVilIDioTV7AKxDZ2dZXm4nJxw";

        private SheetsService _sheetsService;

        static readonly string sheet = "Lista Corretores";

        private int i = 0;

        private int quantidadeCorretores = 0;

        private string urlAtual = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }


        private void inserirNaPlanilha(List<IList<object>> data, object sender, EventArgs e, int i)
        {
            try
            {
                if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
                {

                    Google.Apis.Auth.OAuth2.GoogleCredential credential;
                    // Put your credentials json file in the root of the solution and make sure copy to output dir property is set to always copy 

                    using (var stream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "credentials.json", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {

                        credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

                    }

                    // Create Google Sheets API service.
                    _sheetsService = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = _applicationName
                    });

                    var range = $"{sheet}!A:E";
                    var valueRange = new ValueRange();
                    valueRange.Values = data;

                    var appendRequest = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
                    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var appendReponse = appendRequest.Execute();

                    var html = webBrowser1.Document.Body.InnerHtml;
                    var url = webBrowser1.Url.AbsoluteUri;
                    url = webBrowser1.Url.AbsoluteUri;

                    webBrowser1.AllowNavigation = true;


                    webBrowser1.DocumentCompleted -= ProcessoCorretores;
                    this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(this.button1_Click);
                    this.webBrowser1.Navigate(urlAtual);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void ProcessoCorretores(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            List<IList<object>> dadosCorretor = new List<IList<object>>();
            IList<object> DadosCorretor = new List<object>();

            if ((sender as WebBrowser).ReadyState == System.Windows.Forms.WebBrowserReadyState.Complete)
            {



                // Do what ever you want to do here when page is completely loaded.

                HtmlElementCollection getNomeCorretor = webBrowser1.Document.GetElementsByTagName("h3");





                var url = webBrowser1.Url.AbsoluteUri;

                foreach (HtmlElement elem in getNomeCorretor)
                {
                    //Nome
                    if (elem.InnerHtml != "Telefones" && elem.InnerHtml != null && elem.InnerHtml.Length > 2)
                    {

                        DadosCorretor.Add(elem.InnerHtml);
                    }
                }


                HtmlElementCollection getDadosRestantes = webBrowser1.Document.GetElementsByTagName("span");

                foreach (HtmlElement elem in getDadosRestantes)
                {
                    //Nome
                    if (elem != null && elem.InnerHtml != "" && elem.InnerHtml != null && elem.Parent.InnerHtml != ""
                        && elem.Parent.InnerHtml != null
                        && (elem.Parent.InnerHtml.Contains("CRECI") ||
                        elem.Parent.InnerHtml.Contains("Situação") ||
                        elem.Parent.InnerHtml.Contains("E-Mail Oficial") ||
                        elem.Parent.InnerHtml.Contains("\n                                            <span>")) &&
                        !elem.InnerHtml.Contains("<i class=\"fa fa-spinner fa-spin \"></i>"))
                    {

                        if (DadosCorretor.Contains(" Ativo") && (elem.InnerHtml.Length == 15 || elem.InnerHtml.Length == 16 || elem.InnerHtml.Length == 14))
                        {
                            if (DadosCorretor.Count == 4 || DadosCorretor.Count == 5 || DadosCorretor.Count == 6)
                            {
                                DadosCorretor.Add(elem.InnerHtml);
                            }
                            else
                            {
                                if (DadosCorretor.Contains("("))
                                {
                                    DadosCorretor.Add(elem.InnerHtml);
                                }
                                else
                                {
                                    DadosCorretor.Add("");
                                    DadosCorretor.Add(elem.InnerHtml);
                                }
                            }

                        }
                        else
                        {
                            DadosCorretor.Add(elem.InnerHtml);
                        }



                    }
                }

                dadosCorretor.Add(DadosCorretor);


                inserirNaPlanilha(dadosCorretor, sender, e, i);


            }


        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (webBrowser1.Document != null)
            {
                quantidadeCorretores = webBrowser1.Document.GetElementsByTagName("h6").Count;

                if (quantidadeCorretores > numeroRegistro.Count)
                {


                    if (webBrowser1.ReadyState == System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        this.webBrowser1.DocumentCompleted -= button1_Click;
                        this.webBrowser1.DocumentCompleted += ProcessoCorretores;

                        HtmlElementCollection elems = webBrowser1.Document.GetElementsByTagName("span");
                        var url = webBrowser1.Url.AbsoluteUri;

                        if (webBrowser1.Url.AbsoluteUri == "https://www.crecisp.gov.br/cidadao/listadecorretores")
                        {
                            urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i;
                            url = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i;
                        }

                        foreach (HtmlElement elem in elems)
                        {
                            if (!numeroRegistro.Contains(elem.InnerHtml) && quantidadeCorretores > numeroRegistro.Count)
                            {
                                if (elem.InnerHtml != null && elem.InnerHtml != "" && elem.InnerHtml.Length == 9)
                                {
                                    if (!numeroRegistro.Contains(elem.InnerHtml) && quantidadeCorretores > numeroRegistro.Count)
                                    {

                                        numeroRegistro.Add(elem.InnerHtml);

                                        HtmlElementCollection buttonClick = elem.Parent.Parent.GetElementsByTagName("button");

                                        foreach (HtmlElement elemt in buttonClick)
                                        {
                                            if (elemt.InnerHtml == " Ver Detalhes ")
                                            {
                                                elemt.InvokeMember("click");
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                    else
                                    {
                                        if (webBrowser1.Document.Body.GetAttribute("class") == "navigate")
                                        {
                                            i++;
                                            string[] getPage = null;
                                            url = webBrowser1.Url.AbsoluteUri;
                                            if (url.Contains("firstLetter="))
                                            {
                                                getPage = url.Split('=');

                                                var first = char.Parse(getPage[2]);

                                                numeroRegistro = new List<string>();
                                                webBrowser1.AllowNavigation = true;
                                                this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + first);
                                                urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + first;
                                            }
                                            else
                                            {
                                                numeroRegistro = new List<string>();
                                                webBrowser1.AllowNavigation = true;
                                                this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i);
                                                urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i;
                                            }

                                        }
                                        else
                                        {
                                            url = webBrowser1.Url.AbsoluteUri;
                                            string[] getPage = null;
                                            char first = ' ';
                                            char nextChar = ' ';
                                            i = 0;
                                            if (url.Contains("firstLetter="))
                                            {
                                                getPage = url.Split('=');


                                                if (getPage != null && getPage.Length == 2)
                                                {
                                                    first = char.Parse(getPage[1]);
                                                    nextChar = (char)((char)first + 1);
                                                }

                                                if (getPage != null && getPage.Length == 3)
                                                {
                                                    first = char.Parse(getPage[2]);
                                                    nextChar = (char)((char)first + 1);
                                                }

                                                numeroRegistro = new List<string>();
                                                webBrowser1.AllowNavigation = true;
                                                this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + nextChar);
                                                urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + nextChar;
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {

                    if (webBrowser1.Document.Body.InnerHtml.Contains("navigate"))
                    {
                        i++;
                        string[] getPage = null;
                        var url = webBrowser1.Url.AbsoluteUri;
                        if (url.Contains("firstLetter="))
                        {
                            getPage = url.Split('=');

                            var first = char.Parse(getPage[2]);

                            numeroRegistro = new List<string>();
                            webBrowser1.AllowNavigation = true;
                            this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + first);
                            urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + first;
                        }
                        else
                        {
                            numeroRegistro = new List<string>();
                            webBrowser1.AllowNavigation = true;
                            this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i);
                            urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i;
                        }


                    }
                    else
                    {
                        var url = webBrowser1.Url.AbsoluteUri;
                        string[] getPage = null;
                        char first = ' ';
                        char nextChar = ' ';
                        i = 0;
                        if (url.Contains("firstLetter="))
                        {
                            getPage = url.Split('=');


                            if (getPage != null && getPage.Length == 2)
                            {
                                first = char.Parse(getPage[1]);
                                nextChar = (char)((char)first + 1);
                            }

                            if (getPage != null && getPage.Length == 3)
                            {
                                first = char.Parse(getPage[2]);
                                nextChar = (char)((char)first + 1);
                            }

                            numeroRegistro = new List<string>();
                            webBrowser1.AllowNavigation = true;
                            this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + nextChar);
                            urlAtual = "https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i + "&firstLetter=" + nextChar;
                        }
                    }
                }


            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.webBrowser1.Navigate(textBox1.Text);
        }

        private void Voltar_Click(object sender, EventArgs e)
        {
            (sender as WebBrowser).GoBack();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/buscaporcorretores");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var url = webBrowser1.Url.AbsoluteUri;
            var getPage = url.Split('=');

            webBrowser1.AllowNavigation = true;


            webBrowser1.DocumentCompleted -= button4_Click;
            this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(this.button1_Click);


            if (getPage.Length == 3)
            {
                var pagina = getPage[1].Split('&');
                i = int.Parse(pagina[0]);
                this.webBrowser1.Navigate(url);
                urlAtual = url;
            }

            if (getPage.Length == 2)
            {
                //i = int.Parse(getPage[1]);
                int inteiro;
                var verificacao = Int32.TryParse(getPage[1], out inteiro);

                if (verificacao == false)
                {
                    this.webBrowser1.Navigate(url);
                    urlAtual = url;
                }
                else
                {
                    i = inteiro;
                    this.webBrowser1.Navigate("https://www.crecisp.gov.br/cidadao/listadecorretores?page=" + i);
                    urlAtual = url;
                }
            }

        }
    }
}
