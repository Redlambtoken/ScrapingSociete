using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Text.RegularExpressions;

namespace App
{
    internal class Program
    {
        public static void Main()
        {
            List<string> TabNamePlace = new List<string>();;
            string data = "Nom,Date de Création,Adresse,SIREN,SIRET,CODE APE,rôles";
            // Configurer la licence EPPlus
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Chemin du fichier Excel à lire
            string filePath = @"C:\Users\Jean Fruleux\Downloads\test.xlsx";

            // Ouvrir le fichier Excel
            FileInfo fileInfo = new FileInfo(filePath);
            using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(fileInfo))
            {
                // Sélectionner la première feuille de calcul
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Lire les cellules de la feuille
                int rowCount = worksheet.Dimension.Rows; // Nombre de lignes
                int colCount = worksheet.Dimension.Columns; // Nombre de colomnes
                TabNamePlace = FormatageCellule(rowCount, worksheet);
            }
            //options.AddArgument("headless");
            WebDriver driver = new ChromeDriver();
            string baseUrlSociete = "https://www.societe.com/";
            driver.Navigate().GoToUrl(baseUrlSociete);
            Thread.Sleep(5000);
            data += ScrapingDataCompagnie(TabNamePlace, driver);

            // Assurez-vous que la bibliothèque EPPlus est autorisée à créer des fichiers Excel
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Exemple de chaîne de caractères (données séparées par des virgules)

            // Créer un nouveau fichier Excel
            using (OfficeOpenXml.ExcelPackage packageV1 = new OfficeOpenXml.ExcelPackage())
            {
                // Ajouter une feuille de calcul
                OfficeOpenXml.ExcelWorksheet worksheetV1 = packageV1.Workbook.Worksheets.Add("Feuille1");

                // Diviser les données en lignes
                string[] rows = data.Split('\n');

                for (int i = 0; i < rows.Length; i++)
                {
                    // Diviser chaque ligne en colonnes (par des virgules)
                    string[] columns = rows[i].Split(',');

                    for (int j = 0; j < columns.Length; j++)
                    {
                        // Ajouter les données à la cellule correspondante
                        worksheetV1.Cells[i + 1, j + 1].Value = columns[j];
                    }
                }

                // Définir le chemin où enregistrer le fichier Excel
                string filePathOutput = @"C:\Users\Jean Fruleux\Desktop\outputV1.xlsx";

                // Enregistrer le fichier Excel
                FileInfo excelFile = new FileInfo(filePathOutput);
                packageV1.SaveAs(excelFile);

                Console.WriteLine("Les données ont été écrites dans le fichier Excel avec succès.");
            }
        }


        public static List<string> FormatageCellule(int rowCount, OfficeOpenXml.ExcelWorksheet worksheet)
        {
            List<string> TabNamePlace = new List<string>();
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= 2; col++)
                {
                    // Lire la valeur de chaque Cellules, la mettre dans une liste avec le bon format
                    string stringclean = Regex.Replace((worksheet.Cells[row, col].Value?.ToString()).Normalize(System.Text.NormalizationForm.FormD).ToUpper(), @"[^A-Z0-9\s]", "");
                    TabNamePlace.Add(stringclean);
                }
                Console.WriteLine(); // Sauter à la ligne suivante après avoir lu une ligne entière
            }
            return TabNamePlace;
        }




        public static string ScrapingDataCompagnie(List<string> TabNamePlace, WebDriver driver)
        {
            string DataToAdd = "";
            for (int j = 2; j < TabNamePlace.Count; j += 2)
            {
                var searchInput = driver.FindElement(By.XPath("//*[@placeholder=\"Entreprise, dirigeant, SIREN...\"]"));
                string compagnieToFind = TabNamePlace[j];
                searchInput.SendKeys(compagnieToFind);
                var searchInputButton = driver.FindElement(By.XPath("//*[@id=\"buttsearch\"]"));
                searchInputButton.Click();
                IWebElement divCompagnie;
                try
                {
                    divCompagnie = driver.FindElement(By.XPath("//*[@id=\"result_deno_societe\"]"));
                }
                catch
                {
                    try
                    {
                        divCompagnie = driver.FindElement(By.XPath("//*[@id=\"result_rs_societe\"]"));
                    }
                    catch
                    {
                        continue;
                    }
                }

                IList<IWebElement> aDiv = divCompagnie.FindElements(By.TagName("a"));
                bool temp = false;
                foreach (IWebElement value in aDiv)
                {
                    IList<IWebElement> paragraphs = divCompagnie.FindElements(By.TagName("p"));
                    foreach (IWebElement paragraph in paragraphs)
                    {
                        string paragraphText = Regex.Replace(paragraph.Text, @"[\d\s]", "");
                        Console.WriteLine("-----------------"); //C'est pour afficher les différentes données quand on fait la recherche dans les paragraphes
                        Console.WriteLine(paragraph.Text);
                        Console.WriteLine("-----------------");
                        if (paragraphText == TabNamePlace[j + 1])
                        {
                            value.Click();
                            temp = true;
                            break;
                        }
                    }
                    if (temp)
                    {
                        break;
                    }
                }

                try
                {
                    IWebElement CreateDate = driver.FindElement(By.ClassName("TableTextGenerique"));
                    IWebElement SirenNumberDiv = driver.FindElement(By.XPath("//*[@id=\"siren_number\"]"));
                    IWebElement SirenNumberValue = SirenNumberDiv.FindElement(By.ClassName("copyNumber__copy"));
                    IWebElement SiretNumberDiv = driver.FindElement(By.XPath("//*[@id=\"siret_number\"]"));
                    IWebElement SiretNumberValue = SiretNumberDiv.FindElement(By.ClassName("copyNumber__copy"));
                    IWebElement CodeNAF = driver.FindElement(By.XPath("//*[@id=\"ape-histo-description\"]"));
                    IWebElement DivRenseignement = driver.FindElement(By.XPath("//*[@id=\"rensjur\"]"));
                    IWebElement AdressePostale = DivRenseignement.FindElement(By.CssSelector(".Lien.secondaire"));
                    string[] Tabdirigeant = new string[20];
                    try
                    {
                        IWebElement DivDirigeant = driver.FindElement(By.XPath("//*[@id=\"tabledir\"]"));
                        var childElements = DivDirigeant.FindElements(By.XPath("./*"));
                        int itmp = 0;
                        foreach (IWebElement divChild in childElements)
                        {
                            if (divChild.TagName == "h5")
                            {
                                // Si l'élément est un h5, récupérer le texte
                                string h5Text = divChild.Text;
                                Tabdirigeant[itmp] = h5Text;
                                itmp++;
                            }
                            else if (divChild.TagName == "div")
                            {
                                // Si l'élément est une div, trouver le span à l'intérieur et récupérer son texte
                                IWebElement spanElement = divChild.FindElement(By.TagName("span"));
                                string spanText = spanElement.Text;
                                Tabdirigeant[itmp] = spanText;
                                itmp++;
                            }
                        }
                    }
                    catch
                    {

                    }
                    string Dirigeant = "";
                    for (int i = 0; i < Tabdirigeant.Length; i++)
                    {
                        Console.WriteLine(Tabdirigeant[i]);
                        Dirigeant += Tabdirigeant[i] + " ";
                    }
                    DataToAdd += "\n" + TabNamePlace[j] + "," + CreateDate.Text + "," + AdressePostale.Text + "," + SirenNumberValue.Text + "," + SiretNumberValue.Text + "," + CodeNAF.Text + "," + Dirigeant;

                }
                catch
                {
                    return ""; //je l'ai mis car le code mets une erreur sinon, il ne sert à rien
                }
            }
            return DataToAdd;
        }
    }
}