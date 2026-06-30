namespace Fine
{
    /// <summary>
    /// 违规类型枚举 -- 定义所有支持的违规类型
    /// </summary>
    public enum ViolationTypeEnum
    {
        RouteDeviation,       // 路线偏离
        RedLight,             // 冲红灯
        DriverNoSeatbelt,     // 司机不系安全带
        AttendantNoSeatbelt,  // 照管员不系安全带
        Speeding,             // 超速
        DoorOpen,             // 未关车门
        OffSiteStop,          // 未按站点停靠
        Smoking,              // 车内抽烟
        AttendantPhone        // 照管员玩手机
    }

    /// <summary>
    /// 违规类型配置类 -- 封装每种类型的显示名、模板名、文件夹名、通报文案
    /// </summary>
    public class ViolationTypeConfig
    {
        public ViolationTypeEnum Type { get; set; }
        public string DisplayName { get; set; }
        public string TemplateFileName { get; set; }
        public string FolderSuffix { get; set; }
        public string ReportTitleTemplate { get; set; }
        public string ReportBodyTemplate { get; set; }

        public string OrderPrefix => $"CF-{DisplayName}";

        public string GetTemplatePath(string startupPath)
            => Path.Combine(startupPath, TemplateFileName);
    }

    /// <summary>
    /// 违规类型配置仓库 -- 提供所有类型的配置映射
    /// </summary>
    public static class ViolationTypeRegistry
    {
        private static readonly Dictionary<ViolationTypeEnum, ViolationTypeConfig> _configs;

        static ViolationTypeRegistry()
        {
            _configs = new Dictionary<ViolationTypeEnum, ViolationTypeConfig>
            {
                [ViolationTypeEnum.RouteDeviation] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.RouteDeviation,
                    DisplayName = "路线偏离",
                    TemplateFileName = "Template_RouteDeviation.docx",
                    FolderSuffix = "路线偏离车辆违规处罚通知单",
                    ReportTitleTemplate = "校车路线偏离违规",
                    ReportBodyTemplate = "核实偏离原因（如遇修路、封路等客观原因，请管理员提供证明报备）。\n    再次向车队强调：严禁私自更改交警审批路线。"
                },
                [ViolationTypeEnum.RedLight] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.RedLight,
                    DisplayName = "冲红灯",
                    TemplateFileName = "Template_RedLight.docx",
                    FolderSuffix = "冲红灯车辆违规处罚通知单",
                    ReportTitleTemplate = "校车冲红灯违规",
                    ReportBodyTemplate = "核实冲红灯原因（如紧急避让等特殊情况，请管理员提供证明报备）。\n    再次向车队强调：严禁冲红灯，确保行车安全。"
                },
                [ViolationTypeEnum.DriverNoSeatbelt] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.DriverNoSeatbelt,
                    DisplayName = "司机不系安全带",
                    TemplateFileName = "Template_DriverNoSeatbelt.docx",
                    FolderSuffix = "司机不系安全带车辆违规处罚通知单",
                    ReportTitleTemplate = "校车司机不系安全带违规",
                    ReportBodyTemplate = "核实司机未系安全带原因。\n    再次向车队强调：驾驶员行车必须全程系好安全带。"
                },
                [ViolationTypeEnum.AttendantNoSeatbelt] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.AttendantNoSeatbelt,
                    DisplayName = "照管员不系安全带",
                    TemplateFileName = "Template_AttendantNoSeatbelt.docx",
                    FolderSuffix = "照管员不系安全带车辆违规处罚通知单",
                    ReportTitleTemplate = "校车照管员不系安全带违规",
                    ReportBodyTemplate = "核实照管员未系安全带原因。\n    再次向车队强调：照管员随车必须全程系好安全带。"
                },
                [ViolationTypeEnum.Speeding] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.Speeding,
                    DisplayName = "超速",
                    TemplateFileName = "Template_Speeding.docx",
                    FolderSuffix = "超速车辆违规处罚通知单",
                    ReportTitleTemplate = "校车超速违规",
                    ReportBodyTemplate = "核实超速原因。\n    再次向车队强调：严禁超速行驶，确保学生乘车安全。"
                },
                [ViolationTypeEnum.DoorOpen] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.DoorOpen,
                    DisplayName = "未关车门",
                    TemplateFileName = "Template_DoorOpen.docx",
                    FolderSuffix = "未关车门车辆违规处罚通知单",
                    ReportTitleTemplate = "校车未关车门违规",
                    ReportBodyTemplate = "核实未关车门原因。\n    再次向车队强调：行驶前务必确认车门已完全关闭。"
                },
                [ViolationTypeEnum.OffSiteStop] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.OffSiteStop,
                    DisplayName = "未按站点停靠",
                    TemplateFileName = "Template_OffSiteStop.docx",
                    FolderSuffix = "未按站点停靠车辆违规处罚通知单",
                    ReportTitleTemplate = "校车未按站点停靠违规",
                    ReportBodyTemplate = "核实未按站点停靠原因。\n    再次向车队强调：必须按审批站点停靠上下学生。"
                },
                [ViolationTypeEnum.Smoking] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.Smoking,
                    DisplayName = "车内抽烟",
                    TemplateFileName = "Template_Smoking.docx",
                    FolderSuffix = "车内抽烟车辆违规处罚通知单",
                    ReportTitleTemplate = "校车车内抽烟违规",
                    ReportBodyTemplate = "核实车内抽烟情况。\n    再次向车队强调：校车内严禁吸烟，保障学生健康。"
                },
                [ViolationTypeEnum.AttendantPhone] = new ViolationTypeConfig
                {
                    Type = ViolationTypeEnum.AttendantPhone,
                    DisplayName = "照管员玩手机",
                    TemplateFileName = "Template_AttendantPhone.docx",
                    FolderSuffix = "照管员玩手机车辆违规处罚通知单",
                    ReportTitleTemplate = "校车照管员玩手机违规",
                    ReportBodyTemplate = "核实照管员玩手机情况。\n    再次向车队强调：随车照管员应专注看护学生，严禁玩手机。"
                }
            };
        }

        public static ViolationTypeConfig GetConfig(ViolationTypeEnum type)
            => _configs[type];

        public static List<ViolationTypeEnum> AllTypes => _configs.Keys.ToList();

        public static List<ViolationTypeConfig> AllConfigs => _configs.Values.ToList();
    }
}
