using System.Text.RegularExpressions;

namespace Fine
{
    /// <summary>
    /// 序号管理器 -- 按违规类型独立维护序号
    /// 策略: 扫描目标文件夹已有文件, 推算下一个序号
    /// </summary>
    public static class SequenceNumberManager
    {
        /// <summary>
        /// 从目标文件夹已有的文件中推算下一个序号
        /// 文件名格式: CF-{类型}-{yyyyMMdd}-{seq}{车牌}车辆违规处罚通知单.pdf
        /// </summary>
        public static int GetNextSequence(string targetFolder, ViolationTypeEnum violationType, string dateStr)
        {
            string orderPrefix = $"CF-{ViolationTypeRegistry.GetConfig(violationType).DisplayName}-{dateStr}-";

            if (!Directory.Exists(targetFolder))
                return 1;

            int maxSeq = 0;
            string[] existingFiles = Directory.GetFiles(targetFolder, "*.pdf");

            foreach (string file in existingFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (fileName.StartsWith(orderPrefix))
                {
                    string afterPrefix = fileName.Substring(orderPrefix.Length);
                    Match seqMatch = Regex.Match(afterPrefix, @"^(\d{2})");
                    if (seqMatch.Success && int.TryParse(seqMatch.Groups[1].Value, out int seq))
                    {
                        if (seq > maxSeq) maxSeq = seq;
                    }
                }
            }

            return maxSeq + 1;
        }
    }
}
