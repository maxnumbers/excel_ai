using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace LocalAI.Excel.AddIn
{
    public static class AIFunctions
    {
        private static readonly HttpClient httpClient = new HttpClient() { Timeout = TimeSpan.FromMinutes(2) };
        private static AIConfiguration config = new AIConfiguration();

        #region AI Chat Functions

        [ExcelFunction(
            Name = "AI.CHAT",
            Description = "Generate AI response using Ollama or LM Studio",
            Category = "Local AI",
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_CHAT(
            [ExcelArgument(Name = "prompt", Description = "The prompt to send to AI")] string prompt,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model,
            [ExcelArgument(Name = "temperature", Description = "Response creativity 0-1 (optional, default 0.7)")] object temperature)
        {
            if (string.IsNullOrWhiteSpace(prompt))
                return "Error: No prompt provided";

            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var temp = ExcelHelper.GetOptionalDouble(temperature, 0.7);

            return ExcelAsyncUtil.Run("AI.CHAT", new object[] { prompt, modelName, temp }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, temp).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        [ExcelFunction(
            Name = "AI.SUMMARIZE", 
            Description = "Summarize text using AI",
            Category = "Local AI",
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_SUMMARIZE(
            [ExcelArgument(Name = "text", Description = "Text to summarize")] string text,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model,
            [ExcelArgument(Name = "style", Description = "Summary style: brief, detailed, bullet (optional)")] object style)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Error: No text provided";

            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var summaryStyle = ExcelHelper.GetOptionalString(style, "brief").ToLower();

            string prompt = summaryStyle switch
            {
                "detailed" => $"Please provide a detailed summary of the following text, highlighting key points:\n\n{text}",
                "bullet" => $"Please summarize the following text as bullet points:\n\n{text}",
                _ => $"Please provide a brief summary of the following text:\n\n{text}"
            };

            return ExcelAsyncUtil.Run("AI.SUMMARIZE", new object[] { text, modelName, summaryStyle }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, 0.5).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        [ExcelFunction(
            Name = "AI.TRANSLATE",
            Description = "Translate text using AI", 
            Category = "Local AI",
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_TRANSLATE(
            [ExcelArgument(Name = "text", Description = "Text to translate")] string text,
            [ExcelArgument(Name = "language", Description = "Target language")] string targetLanguage,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model,
            [ExcelArgument(Name = "formal", Description = "Use formal tone (optional)")] object formal)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Error: No text provided";
            if (string.IsNullOrWhiteSpace(targetLanguage))
                return "Error: No target language specified";

            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var useFormal = ExcelHelper.GetOptionalBool(formal, false);
            var tone = useFormal ? "formal" : "natural";

            var prompt = $"Translate the following text to {targetLanguage} using a {tone} tone. Only return the translation:\n\n{text}";

            return ExcelAsyncUtil.Run("AI.TRANSLATE", new object[] { text, targetLanguage, modelName }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, 0.3).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        [ExcelFunction(
            Name = "AI.ANALYZE",
            Description = "Analyze data and answer questions using AI",
            Category = "Local AI", 
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_ANALYZE(
            [ExcelArgument(Name = "data_range", Description = "Range of data to analyze")] object[,] dataRange,
            [ExcelArgument(Name = "question", Description = "Question about the data")] string question,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model)
        {
            if (dataRange == null || dataRange.GetLength(0) == 0)
                return "Error: No data provided";
            if (string.IsNullOrWhiteSpace(question))
                return "Error: No question provided";

            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var dataText = ExcelHelper.ConvertRangeToText(dataRange);
            var prompt = $"Data to analyze:\n{dataText}\n\nQuestion: {question}\n\nPlease analyze this data and answer the question with specific references to the data.";

            return ExcelAsyncUtil.Run("AI.ANALYZE", new object[] { dataRange, question, modelName }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, 0.4).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        [ExcelFunction(
            Name = "AI.SENTIMENT",
            Description = "Analyze sentiment of text using AI",
            Category = "Local AI",
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_SENTIMENT(
            [ExcelArgument(Name = "text", Description = "Text to analyze")] string text,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model,
            [ExcelArgument(Name = "detailed", Description = "Return detailed analysis (optional)")] object detailed)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Error: No text provided";

            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var useDetailed = ExcelHelper.GetOptionalBool(detailed, false);

            var prompt = useDetailed
                ? $"Analyze the sentiment of the following text and provide a detailed explanation with confidence scores:\n\n{text}"
                : $"Analyze the sentiment of the following text. Return only: Positive, Negative, or Neutral:\n\n{text}";

            return ExcelAsyncUtil.Run("AI.SENTIMENT", new object[] { text, modelName }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, 0.2).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        [ExcelFunction(
            Name = "AI.CODE",
            Description = "Generate code or Excel formulas using AI",
            Category = "Local AI",
            IsThreadSafe = true,
            IsVolatile = false
        )]
        public static object AI_CODE(
            [ExcelArgument(Name = "description", Description = "Description of what code/formula you need")] string description,
            [ExcelArgument(Name = "language", Description = "Programming language or 'excel' for formulas (optional)")] object language,
            [ExcelArgument(Name = "model", Description = "AI model to use (optional)")] object model)
        {
            if (string.IsNullOrWhiteSpace(description))
                return "Error: No description provided";

            var lang = ExcelHelper.GetOptionalString(language, "excel");
            var modelName = ExcelHelper.GetOptionalString(model, config.DefaultModel);
            var prompt = $"Generate {lang} code for: {description}\n\nOnly return the code, no explanations unless specifically requested.";

            return ExcelAsyncUtil.Run("AI.CODE", new object[] { description, lang, modelName }, () =>
            {
                try
                {
                    return GenerateAIResponse(prompt, modelName, 0.3).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            });
        }

        #endregion

        #region Core AI Communication

        private static async Task<string> GenerateAIResponse(string prompt, string model, double temperature)
        {
            try
            {
                if (config.Service == "lmstudio" || config.ApiUrl.Contains("1234"))
                {
                    return await CallLMStudioAPI(prompt, model, temperature);
                }
                else
                {
                    return await CallOllamaAPI(prompt, model, temperature);
                }
            }
            catch (HttpRequestException)
            {
                return $"Connection Error: Make sure {config.Service} is running on {config.ApiUrl}";
            }
            catch (TaskCanceledException)
            {
                return "Timeout: AI request took too long. Try a simpler prompt or check your AI service.";
            }
        }

        private static async Task<string> CallOllamaAPI(string prompt, string model, double temperature)
        {
            var request = new
            {
                model = model,
                prompt = prompt,
                stream = false,
                options = new { temperature = temperature }
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync($"{config.ApiUrl}/api/generate", content);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JsonConvert.DeserializeObject<OllamaResponse>(responseJson);

            return result?.Response ?? "No response received from Ollama";
        }

        private static async Task<string> CallLMStudioAPI(string prompt, string model, double temperature)
        {
            var request = new
            {
                model = model,
                messages = new[] { new { role = "user", content = prompt } },
                temperature = temperature,
                max_tokens = 1000
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync($"{config.ApiUrl}/v1/chat/completions", content);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JsonConvert.DeserializeObject<LMStudioResponse>(responseJson);

            return result?.Choices?[0]?.Message?.Content ?? "No response received from LM Studio";
        }

        #endregion

        #region Ribbon Commands

        [ExcelCommand(MenuName = "Local AI", MenuText = "Show Configuration")]
        public static void ShowConfigurationPane()
        {
            var form = new ConfigurationForm(config);
            form.ShowDialog();
        }

        [ExcelCommand(MenuName = "Local AI", MenuText = "Test Connection")]
        public static void TestAIConnection()
        {
            Task.Run(async () =>
            {
                try
                {
                    var result = await GenerateAIResponse("Test connection", config.DefaultModel, 0.7);
                    MessageBox.Show($"Connection successful!\nResponse: {result.Substring(0, Math.Min(100, result.Length))}...", 
                        "AI Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Connection failed: {ex.Message}", "AI Connection Test", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
        }

        [ExcelCommand(MenuName = "Local AI", MenuText = "Function Help")]
        public static void ShowHelpDialog()
        {
            var help = @"Local AI Excel Functions:

=AI.CHAT(prompt, [model], [temperature])
  General AI conversation and responses

=AI.SUMMARIZE(text, [model], [style])  
  Summarize text (style: brief, detailed, bullet)

=AI.TRANSLATE(text, language, [model], [formal])
  Translate to any language

=AI.ANALYZE(data_range, question, [model])
  Analyze Excel data and answer questions

=AI.SENTIMENT(text, [model], [detailed])
  Analyze text sentiment 

=AI.CODE(description, [language], [model])
  Generate code or Excel formulas

All functions work asynchronously and won't freeze Excel.
Configure your AI service using the ribbon button.";

            MessageBox.Show(help, "Local AI Function Guide", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }
}