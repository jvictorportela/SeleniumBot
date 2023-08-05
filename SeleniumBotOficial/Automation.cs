using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumBotOficial
{
    internal class Automation
    {
        public IWebDriver driver;
        public string nameCidade;
        public string email;
        public string endereco;
        public string cep;
        public string uf;

        public Automation()
        {
            driver = new ChromeDriver();
        }

        public void StartWeb()
        {
            driver.Navigate().GoToUrl("https://pt.fakenamegenerator.com/gen-random-br-br.php");
            driver.Manage().Window.Maximize();

            //Time Sleep adicionado, pois o site demora a liberar o elemento.
            Thread.Sleep(2000);
            Console.WriteLine("Após 2 segundos");

            IWebElement primeiroElemento = driver.FindElement(By.XPath("//*[@id='gen']/option[1]"));
            primeiroElemento.Click();
            Console.WriteLine("Sexo definido!");

            Thread.Sleep(2000);

            IWebElement segundoElemento = driver.FindElement(By.XPath("//*[@id=\'n\']/option[4]"));
            segundoElemento.Click();
            Console.WriteLine("Conjunto de nome definido!");

            IWebElement terceiroElemento = driver.FindElement(By.XPath("//*[@id=\'c\']/option[3]"));
            terceiroElemento.Click();
            Console.WriteLine("País definido!");

            Thread.Sleep(2000);
            Console.WriteLine("Após 2 segundos");

            driver.FindElement(By.XPath("//*[@id=\'genbtn\']")).Click();
            Console.WriteLine("Gerou!");

            Thread.Sleep(2000);
            Console.WriteLine("Após 2 segundos");

            string nomeCompleto = driver.FindElement(By.XPath("//*[@id=\'details\']/div[2]/div[2]/div/div[1]/h3")).Text;
            Console.WriteLine("Nome capturado: " + nomeCompleto);

            IWebElement divCity = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.address > div"));
            Thread.Sleep(3000);
            string divCidade = divCity.Text;
            string[] cidade = divCidade.Split('\n');

            if (cidade.Length >= 3)
            {
                nameCidade = cidade[1].Trim();
                Console.WriteLine("Cidade econtrada: " + nameCidade);
            }

            else
            {
                Console.WriteLine("Não foi possível encontrar a segunda linha.");
            }

            Thread.Sleep(1000);

            IWebElement divPais = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.extra > dl:nth-child(6) > dd"));
            Thread.Sleep(2000);
            string elementoPais = divPais.Text;
            if (elementoPais == "55")
            {
                elementoPais = "Brasil";
                Console.WriteLine("País capturado: " + elementoPais);
            }
            else
            {
                elementoPais = "País desconhecido";
            }

            Thread.Sleep(2000);

            IWebElement containerEmail = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.extra > dl:nth-child(12) > dd"));
            string realEmail = containerEmail.Text;
            int atIndex = realEmail.IndexOf(".com"); //Encontra a posição do .com
            int atIndex1 = realEmail.IndexOf(".br");
            
            if (atIndex != -1)
            {
                email = realEmail.Substring(0, atIndex + 4); //Captura a parte até o .com
                Console.WriteLine("O elemento contém um email válido: " + email);
            }

            else if (atIndex1 != -1)
            {
                email = realEmail.Substring(0, atIndex1 + 3); //Captura a parte até o .br
                Console.WriteLine("O elemento contém um email válido: " + email);
            }

            else
            {
                Console.WriteLine("O elemento não contém um email válido.");
            }

            Thread.Sleep(2000);

            IWebElement elementEndereco = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.address > div"));
            Thread.Sleep(2000);

            string[] enderecoLines = elementEndereco.Text.Split('\n');

            if (enderecoLines.Length >= 1)
            {
                endereco = enderecoLines[0].Trim();
                Console.WriteLine("Endereço capturado: " + endereco);
            }

            else
            {
                Console.WriteLine("Não foi possível localizar esse endereço.");
            }

            Thread.Sleep(2000);

            IWebElement elementCep = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.address > div"));
            Thread.Sleep(3000);
            string[] cepLines = elementCep.Text.Split('\n');

            if (cepLines.Length >= 1)
            {
                cep = cepLines[2].Trim();
                Console.WriteLine("CEP capturado: " + cep);
            }

            else
            {
                Console.WriteLine("CEP não econtrado.");
            }

            Thread.Sleep(2000);

            IWebElement elementUf = driver.FindElement(By.CssSelector("#details > div.content > div.info > div > div.address > div"));
            Thread.Sleep(2000);
            string[] ufLines = elementUf.Text.Split('\n');

            if (ufLines.Length >= 2)
            {
                string segundaLinha = ufLines[1].Trim();

                if (segundaLinha.Length >= 2)
                {
                    uf = segundaLinha.Substring(segundaLinha.Length - 2);
                    Console.WriteLine("UF encontrada: " + uf);
                }
                else
                {
                    Console.WriteLine("Nada encontrado.");
                }
            }
            else
            {
                Console.WriteLine("UF não encontrada.");
            }


            Thread.Sleep(2000);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string filePath = "dados_gerados.xlsx";

            using (ExcelPackage package  = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Dados gerados");

                //Cabeçalhos
                worksheet.Cells["A1"].Value = "Nome Completo;";
                worksheet.Cells["B1"].Value = "Cidade";
                worksheet.Cells["C1"].Value = "País";
                worksheet.Cells["D1"].Value = "Email";
                worksheet.Cells["E1"].Value = "Endereço";
                worksheet.Cells["F1"].Value = "CEP";
                worksheet.Cells["G1"].Value = "UF";

                //Conteúdos
                worksheet.Cells["A2"].Value = nomeCompleto;
                worksheet.Cells["B2"].Value = nameCidade;
                worksheet.Cells["C2"].Value = elementoPais;
                worksheet.Cells["D2"].Value = email;
                worksheet.Cells["E2"].Value = endereco;
                worksheet.Cells["F2"].Value = cep;
                worksheet.Cells["G2"].Value = uf;


                FileInfo fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }

            Console.WriteLine("Planilha gerada!");
        }
    }
}
