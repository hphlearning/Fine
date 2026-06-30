using System.Text.RegularExpressions;

namespace Fine
{
    /// <summary>
    /// OCR文本解析结果
    /// </summary>
    public class OcrParseResult
    {
        public string Plate { get; set; } = "";
        public string Time { get; set; } = "";
        public bool IsFullyExtracted => !string.IsNullOrEmpty(Plate) && !string.IsNullOrEmpty(Time);
    }

    /// <summary>
    /// OCR文本解析器 -- 从剪贴板OCR文本中提取车牌和违规时间
    /// </summary>
    public static class OcrParser
    {
        private static readonly Regex PlateRegex = new(@"(粤[A-Z][A-Z0-9]{4,5})", RegexOptions.Compiled);
        private static readonly Regex TimeRegex = new(@"(\d{4}-\d{2}-\d{2})\s?(\d{2}:\d{2}:\d{2})", RegexOptions.Compiled);

        /// <summary>
        /// 从原始OCR文本中提取车牌和违规时间
        /// </summary>
        public static OcrParseResult Parse(string rawText)
        {
            var result = new OcrParseResult();

            Match plateMatch = PlateRegex.Match(rawText);
            if (plateMatch.Success)
                result.Plate = plateMatch.Groups[1].Value;

            Match timeMatch = TimeRegex.Match(rawText);
            if (timeMatch.Success)
                result.Time = timeMatch.Groups[1].Value + " " + timeMatch.Groups[2].Value;

            return result;
        }
    }
}
