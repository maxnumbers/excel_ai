using System;
using System.Configuration;
using System.IO;

namespace LocalAI.Excel.AddIn
{
    public class AIConfiguration
    {
        private const string CONFIG_FILE = "LocalAI.config";

        public string Service { get; set; } = "ollama";
        public string ApiUrl { get; set; } = "http://localhost:11434";
        public string DefaultModel { get; set; } = "llama3";

        public AIConfiguration()
        {
            LoadConfiguration();
        }

        public void LoadConfiguration()
        {
            try
            {
                var configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LocalAI-Excel", CONFIG_FILE);

                if (File.Exists(configPath))
                {
                    var lines = File.ReadAllLines(configPath);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue;

                        var parts = line.Split('=');
                        if (parts.Length == 2)
                        {
                            var key = parts[0].Trim();
                            var value = parts[1].Trim();

                            switch (key.ToLower())
                            {
                                case "service":
                                    Service = value;
                                    break;
                                case "apiurl":
                                    ApiUrl = value;
                                    break;
                                case "defaultmodel":
                                    DefaultModel = value;
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    // Create default config
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading configuration: {ex.Message}");
            }
        }

        public void SaveConfiguration()
        {
            try
            {
                var configDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LocalAI-Excel");
                Directory.CreateDirectory(configDir);

                var configPath = Path.Combine(configDir, CONFIG_FILE);

                var configContent = $@"# Local AI Excel Add-in Configuration
# Service: ollama or lmstudio
Service={Service}

# API URL for your AI service
ApiUrl={ApiUrl}

# Default model to use
DefaultModel={DefaultModel}";

                File.WriteAllText(configPath, configContent);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving configuration: {ex.Message}");
            }
        }
    }
}