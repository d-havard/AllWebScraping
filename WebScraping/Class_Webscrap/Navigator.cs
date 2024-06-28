using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Playwright;

namespace Class_Webscrap
{
    public class Navigator
    {

        List<string> jsonResponses = new List<string>();

        /// <summary>
        /// Launch the emulation of the navigator and catch the JSON responses to save them in different files
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<List<string>> LaunchNavigatorProcess(string url)
        {
            // Initialiser Playwright
            using var playwright = await Playwright.CreateAsync();
            var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = true });

            var page = await browser.NewPageAsync();

            // Liste pour stocker les réponses JSON
            jsonResponses = new List<string>();

            page.Response += async (sender, response) =>
            {
                try
                {
                    // Vérifier si le content-type est JSON
                    if (response.Headers["content-type"] != null && response.Headers["content-type"].Contains("application/json"))
                    {
                        // Lire le contenu de la réponse
                        var jsonResponse = await response.TextAsync();
                        jsonResponses.Add(jsonResponse);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de l'interception de la réponse : {ex.Message}");
                }
            };

            await page.GotoAsync(url);

            await page.WaitForTimeoutAsync(5000);

            int count = 1;
            foreach (var jsonResponse in jsonResponses)
            {
                var fileName = $"response_{count++}.json";
                await File.WriteAllTextAsync(fileName, jsonResponse);
                Console.WriteLine($"Enregistré : {fileName}");
            }

            // Fermer le navigateur
            await browser.CloseAsync();

            return jsonResponses;
        }
    }
}
