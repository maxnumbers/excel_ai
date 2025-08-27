using System;
using System.Drawing;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace LocalAI.Excel.AddIn
{
    public partial class ConfigurationForm : Form
    {
        private AIConfiguration config;
        private static readonly HttpClient httpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(10) };

        public ConfigurationForm(AIConfiguration configuration)
        {
            config = configuration;
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "Local AI Configuration";
            this.Size = new Size(450, 300);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // Service selection
            var serviceLabel = new Label()
            {
                Text = "AI Service:",
                Location = new Point(20, 20),
                Size = new Size(100, 20)
            };
            this.Controls.Add(serviceLabel);

            var serviceCombo = new ComboBox()
            {
                Name = "serviceCombo",
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(130, 18),
                Size = new Size(150, 25)
            };
            serviceCombo.Items.AddRange(new[] { "ollama", "lmstudio" });
            this.Controls.Add(serviceCombo);

            // API URL
            var urlLabel = new Label()
            {
                Text = "API URL:",
                Location = new Point(20, 55),
                Size = new Size(100, 20)
            };
            this.Controls.Add(urlLabel);

            var urlTextBox = new TextBox()
            {
                Name = "urlTextBox",
                Location = new Point(130, 53),
                Size = new Size(250, 25)
            };
            this.Controls.Add(urlTextBox);

            // Default Model
            var modelLabel = new Label()
            {
                Text = "Default Model:",
                Location = new Point(20, 90),
                Size = new Size(100, 20)
            };
            this.Controls.Add(modelLabel);

            var modelTextBox = new TextBox()
            {
                Name = "modelTextBox",
                Location = new Point(130, 88),
                Size = new Size(150, 25)
            };
            this.Controls.Add(modelTextBox);

            // Test button
            var testButton = new Button()
            {
                Name = "testButton",
                Text = "Test Connection",
                Location = new Point(290, 86),
                Size = new Size(100, 30)
            };
            testButton.Click += TestConnection_Click;
            this.Controls.Add(testButton);

            // Status label
            var statusLabel = new Label()
            {
                Name = "statusLabel",
                Text = "Ready",
                Location = new Point(20, 130),
                Size = new Size(370, 60),
                ForeColor = Color.Blue
            };
            this.Controls.Add(statusLabel);

            // Buttons
            var okButton = new Button()
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Location = new Point(220, 220),
                Size = new Size(80, 30)
            };
            okButton.Click += Save_Click;
            this.Controls.Add(okButton);

            var cancelButton = new Button()
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(310, 220),
                Size = new Size(80, 30)
            };
            this.Controls.Add(cancelButton);

            // Service change handler
            serviceCombo.SelectedIndexChanged += (sender, e) =>
            {
                var service = serviceCombo.SelectedItem.ToString();
                if (service == "ollama")
                    urlTextBox.Text = "http://localhost:11434";
                else if (service == "lmstudio")
                    urlTextBox.Text = "http://localhost:1234";
            };
        }

        private void LoadCurrentSettings()
        {
            var serviceCombo = (ComboBox)this.Controls["serviceCombo"];
            var urlTextBox = (TextBox)this.Controls["urlTextBox"];
            var modelTextBox = (TextBox)this.Controls["modelTextBox"];

            serviceCombo.SelectedItem = config.Service;
            urlTextBox.Text = config.ApiUrl;
            modelTextBox.Text = config.DefaultModel;
        }

        private async void TestConnection_Click(object sender, EventArgs e)
        {
            var statusLabel = (Label)this.Controls["statusLabel"];
            var urlTextBox = (TextBox)this.Controls["urlTextBox"];
            var serviceCombo = (ComboBox)this.Controls["serviceCombo"];

            statusLabel.Text = "Testing connection...";
            statusLabel.ForeColor = Color.Blue;

            try
            {
                var testUrl = serviceCombo.SelectedItem.ToString() == "ollama" 
                    ? $"{urlTextBox.Text}/api/tags"
                    : $"{urlTextBox.Text}/v1/models";

                var response = await httpClient.GetAsync(testUrl);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                dynamic result = JsonConvert.DeserializeObject(content);

                int modelCount = 0;
                if (serviceCombo.SelectedItem.ToString() == "ollama")
                    modelCount = result?.models?.Count ?? 0;
                else
                    modelCount = result?.data?.Count ?? 0;

                statusLabel.Text = $"✓ Connected successfully! Found {modelCount} model(s).";
                statusLabel.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"✗ Connection failed: {ex.Message}";
                statusLabel.ForeColor = Color.Red;
            }
        }

        private void Save_Click(object sender, EventArgs e)
        {
            var serviceCombo = (ComboBox)this.Controls["serviceCombo"];
            var urlTextBox = (TextBox)this.Controls["urlTextBox"];
            var modelTextBox = (TextBox)this.Controls["modelTextBox"];

            config.Service = serviceCombo.SelectedItem?.ToString() ?? "ollama";
            config.ApiUrl = urlTextBox.Text;
            config.DefaultModel = modelTextBox.Text;

            config.SaveConfiguration();

            MessageBox.Show("Configuration saved successfully!", "Local AI Configuration", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}