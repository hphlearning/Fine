namespace Fine
{
    /// <summary>
    /// 模板加载器 -- 按违规类型查找并加载对应Word模板
    /// </summary>
    public static class TemplateLoader
    {
        /// <summary>
        /// 获取指定违规类型的模板文件路径
        /// 如果专用模板不存在, 则回退到通用 Template.docx
        /// </summary>
        public static string GetTemplatePath(ViolationTypeEnum violationType, string startupPath)
        {
            var config = ViolationTypeRegistry.GetConfig(violationType);
            string specificTemplate = config.GetTemplatePath(startupPath);

            if (File.Exists(specificTemplate))
                return specificTemplate;

            // 回退: 使用通用模板
            string fallbackTemplate = Path.Combine(startupPath, "Template.docx");
            if (File.Exists(fallbackTemplate))
            {
                MessageBox.Show(
                    $"未找到专用模板【{config.TemplateFileName}】，已自动回退使用通用模板 Template.docx。\n\n建议：请复制 Template.docx 并重命名为 {config.TemplateFileName}，放在程序目录下。",
                    "模板回退提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return fallbackTemplate;
            }

            // 两种模板都没有
            MessageBox.Show(
                "程序目录下未找到任何可用模板！\n请确保至少存在 Template.docx。",
                "找不到模板",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return "";
        }

        /// <summary>
        /// 检查所有违规类型的模板是否齐全, 返回缺失列表
        /// </summary>
        public static List<string> GetMissingTemplates(string startupPath)
        {
            var missing = new List<string>();
            foreach (var config in ViolationTypeRegistry.AllConfigs)
            {
                if (!File.Exists(config.GetTemplatePath(startupPath)))
                    missing.Add(config.TemplateFileName);
            }
            return missing;
        }
    }
}
