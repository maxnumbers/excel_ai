using Newtonsoft.Json;

namespace LocalAI.Excel.AddIn
{
    public class OllamaResponse
    {
        [JsonProperty("response")]
        public string Response { get; set; }

        [JsonProperty("done")]
        public bool Done { get; set; }

        [JsonProperty("error")]
        public string Error { get; set; }
    }

    public class LMStudioResponse
    {
        [JsonProperty("choices")]
        public LMStudioChoice[] Choices { get; set; }

        [JsonProperty("error")]
        public LMStudioError Error { get; set; }
    }

    public class LMStudioChoice
    {
        [JsonProperty("message")]
        public LMStudioMessage Message { get; set; }

        [JsonProperty("finish_reason")]
        public string FinishReason { get; set; }
    }

    public class LMStudioMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }
    }

    public class LMStudioError
    {
        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }
    }
}