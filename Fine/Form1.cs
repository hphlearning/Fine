using ExcelDataReader;
using IWshRuntimeLibrary;
using System;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Microsoft.VisualBasic;
using System.IO;
using System.Data;
using ExcelDataReader;

namespace Fine
{
    public partial class Form1 : Form
    {

        // === 补充一个私有变量，用来存储当前预览的图片对象 ===
        // 把之前的 private Image currentPreviewImage; 改成下面这样：
        private System.Drawing.Image currentPreviewImage;

        public Form1()
        {
            InitializeComponent();
            // 开启按键预览，允许窗体捕捉 Shift + Enter
            this.KeyPreview = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Clipboard.ContainsText())
            {
                MessageBox.Show("剪贴板中没有文字！请先用微信截图识别并复制文字。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string clipboardText = Clipboard.GetText();

            // 默认清空
            txtPlate.Text = "";
            txtTime.Text = "";

            // 正则匹配车牌
            Match plateMatch = Regex.Match(clipboardText, @"(粤[A-Z][A-Z0-9]{4,5})");
            if (plateMatch.Success)
            {
                txtPlate.Text = plateMatch.Groups[1].Value;
            }

            // 正则匹配时间 (兼容微信识别可能少空格的情况)
            // 升级版正则匹配：用两个括号 () 分别抓住“日期”和“时间”
            Match timeMatch = Regex.Match(clipboardText, @"(\d{4}-\d{2}-\d{2})\s?(\d{2}:\d{2}:\d{2})");
            if (timeMatch.Success)
            {
                // 强制在日期和时间中间补上一个空格，完美还原标准格式
                txtTime.Text = timeMatch.Groups[1].Value + " " + timeMatch.Groups[2].Value;
            }

            if (string.IsNullOrEmpty(txtPlate.Text) || string.IsNullOrEmpty(txtTime.Text))
            {
                MessageBox.Show("未能完全提取出车牌或时间，请在左侧文本框手动补充修改！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void btnLoadImagePreview_Click(object sender, EventArgs e)
        {
            // 1. 检查剪贴板里有没有图片
            if (Clipboard.ContainsImage())
            {
                // 2. [关键] 内存管理：如果预览框里已经有一张旧图，必须手动释放它！
                // 否则程序运行一整天会慢慢变慢。
                if (currentPreviewImage != null)
                {
                    currentPreviewImage.Dispose(); // 释放旧图片内存
                    picPreview.Image = null;       // 清空控件显示
                }

                // 3. 从剪贴板拿走图片并显示
                // 把之前的 Image originalImage = Clipboard.GetImage(); 改成：
                System.Drawing.Image originalImage = Clipboard.GetImage();
                currentPreviewImage = originalImage; // 存入变量管理
                picPreview.Image = currentPreviewImage; // 填入控件显示
            }
            else
            {
                MessageBox.Show("剪贴板中没有图片！请先用微信按 Alt+A 截图并复制图片证据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ==========================================
        // [新增] 窗口关闭时，确保释放最后一张图片的内存
        // ==========================================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (currentPreviewImage != null)
            {
                currentPreviewImage.Dispose();
            }
            base.OnFormClosing(e);
        }


        // 监听全局快捷键 Shift + Enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Shift | Keys.Enter))
            {
                GenerateTicket();
                return true; // 表示按键已处理，不往下传递
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // 核心方法：结合文本框内容 + 剪贴板图片，生成 Word
        private void GenerateTicket()
        {
            // 1. 基础数据准备
            string templatePath = Path.Combine(Application.StartupPath, "Template.docx");
            string carPlate = txtPlate.Text.Trim();
            string penaltyTime = txtTime.Text.Trim();
            string seq = txtSeq.Text.Trim();

            if (!DateTime.TryParse(penaltyTime, out DateTime alarmDateTime))
            {
                MessageBox.Show("时间格式不正确！");
                return;
            }

            // 2. 构造单号与路径
            // 2. 构造单号与路径 (升级版：自动创建每日文件夹)
            string orderNo = $"CF-{alarmDateTime.ToString("yyyyMMdd")}-{seq}";
            string fileNameNoExt = $"{orderNo}{carPlate}车辆违规处罚通知单";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // 从报警时间里提取出类似 "20260316" 的字符串
            string dateStr = alarmDateTime.ToString("yyyyMMdd");

            // 拼出当天的文件夹名字，例如：CF-20260316车辆违规处罚通知单
            string folderName = $"CF-{dateStr}车辆违规处罚通知单";
            string targetFolder = Path.Combine(desktopPath, folderName);

            // 如果桌面没有这个日期的文件夹，程序就自动建一个！
            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            // 最终的保存路径 (把生成的 Word 和 PDF 都放进这个专属文件夹里)
            string wordPath = Path.Combine(targetFolder, fileNameNoExt + ".docx");
            string pdfPath = Path.Combine(targetFolder, fileNameNoExt + ".pdf");

            try
            {
                // --- 第一步：生成 Word ---
                using (DocX document = DocX.Load(templatePath))
                {
                    document.ReplaceText("{{单号}}", orderNo ?? "");
                    document.ReplaceText("{{签发日期}}", alarmDateTime.ToString("yyyy年M月d日"));
                    document.ReplaceText("{{车牌号码}}", carPlate ?? "");
                    document.ReplaceText("{{违规时间}}", alarmDateTime.ToString("yyyy年M月d日 HH:mm"));

                    // 插入预览框图片
                    if (picPreview.Image != null)
                    {
                        string tempPath = Path.Combine(Path.GetTempPath(), "temp_violation.jpg");
                        picPreview.Image.Save(tempPath, System.Drawing.Imaging.ImageFormat.Jpeg);

                        Xceed.Document.NET.Image wordImage = document.AddImage(tempPath);
                        var picture = wordImage.CreatePicture();
                        picture.Width = 500;
                        picture.Height = 400;

                        document.InsertParagraph().AppendPicture(picture);
                    }

                    // 保存 Word 文件
                    document.SaveAs(wordPath);
                } // 离开这个大括号，Word 文件就会被自动关闭释放

                // --- 第二步：立即自动转 PDF ---
                Spire.Doc.Document spireDoc = new Spire.Doc.Document();
                spireDoc.LoadFromFile(wordPath);
                spireDoc.SaveToFile(pdfPath, Spire.Doc.FileFormat.PDF);
                spireDoc.Close();

                // 3. 自动化处理：单号结尾自动 +1
                if (int.TryParse(seq, out int currentNumber))
                {
                    txtSeq.Text = (currentNumber + 1).ToString("D2");
                }

                // 4. 汇总提示
                MessageBox.Show($"处理成功！\n\n已在桌面生成：\n1. Word 存档\n2. PDF 处罚单\n\n单号：{orderNo}",
                                "全自动处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理过程中发生错误：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }







        private void button2_Click(object sender, EventArgs e)
        {
            // 1. 创建一个“选择文件”对话框
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 设置对话框的标题和过滤条件（只允许选择 .docx 文件）
                openFileDialog.Title = "请选择你的 Word 模板文件";
                openFileDialog.Filter = "Word 文档 (*.docx)|*.docx";
                openFileDialog.Multiselect = false; // 不允许一次选多个

                // 2. 如果用户在对话框里点击了“确定”按钮
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // 获取用户选中的那个文件的完整路径
                        string sourceFilePath = openFileDialog.FileName;

                        // 定义我们要把它复制到的目标路径：即当前程序运行目录下的 Template.docx
                        string targetFilePath = Path.Combine(Application.StartupPath, "Template.docx");

                        // 3. 执行复制！
                        // (第三个参数 true 非常关键，意思是如果运行目录里已经有一个旧的 Template.docx 了，就直接覆盖掉它)
                        System.IO.File.Copy(sourceFilePath, targetFilePath, true);

                        // 4. 搞定，弹窗提示
                        MessageBox.Show("模板导入成功！\n程序现在将使用你最新导入的模板。", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        // 万一因为权限或者文件被占用报错了，捕获错误并提示
                        MessageBox.Show("导入模板时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 1. 创建选择文件对话框
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "请选择要批量转换的 Word 处罚单";
                openFileDialog.Filter = "Word 文档 (*.docx)|*.docx";

                // 【核心修改】开启多选功能
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    int successCount = 0;
                    int failCount = 0;

                    // 2. 遍历每一个选中的文件路径
                    foreach (string sourceWordPath in openFileDialog.FileNames)
                    {
                        // 自动生成 PDF 的保存路径（同目录下同名）
                        string targetPdfPath = sourceWordPath.Replace(".docx", ".pdf");

                        try
                        {
                            // 3. 执行转换逻辑
                            Spire.Doc.Document spireDoc = new Spire.Doc.Document();
                            spireDoc.LoadFromFile(sourceWordPath);
                            spireDoc.SaveToFile(targetPdfPath, Spire.Doc.FileFormat.PDF);

                            successCount++;
                        }
                        catch (Exception)
                        {
                            // 如果某个文件转换失败（比如文件正被打开），记录下来，不中断后续转换
                            failCount++;
                        }
                    }

                    // 4. 转换完成后给出一个汇总提示
                    MessageBox.Show($"批量处理完成！\n\n? 成功：{successCount} 个\n? 失败：{failCount} 个",
                                    "转换结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            GenerateTicket(); // 调用生成 Word 的核心逻辑
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        // 定义一个变量，记录上一次抓到的图片，防止重复刷新导致闪烁
        private System.Drawing.Image _lastImage;

        private void timer1_Tick(object sender, EventArgs e)
        {
            // 1. 先判断剪贴板里到底有没有图，没有就直接跳过
            if (!Clipboard.ContainsImage())
            {
                return;
            }

            try
            {
                // 2. 获取图片
                System.Drawing.Image img = Clipboard.GetImage();
                if (img == null) return;

                // 3. 只有当图片真的变了，才更新预览框，防止一直刷新导致闪烁
                // 注意：简单起见，我们直接对比图片的长宽。如果长宽变了或者还没抓过图，就更新。
                if (_lastImage == null || img.Width != _lastImage.Width || img.Height != _lastImage.Height)
                {
                    _lastImage?.Dispose(); // 释放旧内存
                    _lastImage = img;
                    picPreview.Image = _lastImage;
                }
            }
            catch
            {
                // 剪贴板偶尔会被其他程序占用导致读取失败，这里加个保护防止软件崩溃
            }
        }

        private void btnShortcut_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 获取桌面路径
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                // 2. 设置快捷方式的名字 (改成你喜欢的名字)
                string shortcutPath = Path.Combine(desktopPath, "校车违章生成工具.lnk");

                // 3. 获取当前程序 (.exe) 的完整路径
                string targetPath = Application.ExecutablePath;

                // 4. 创建快捷方式对象
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

                // 5. 配置快捷方式属性
                shortcut.TargetPath = targetPath; // 程序路径
                shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath); // 运行目录
                shortcut.WindowStyle = 1; // 普通窗口模式
                shortcut.Description = "东莞校车违章自动生成系统"; // 鼠标悬停时的描述
                shortcut.IconLocation = targetPath + ",0"; // 使用程序自带的图标

                // 6. 保存到桌面
                shortcut.Save();

                MessageBox.Show("桌面快捷方式创建成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("创建失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 1. 弹出一个窗口，让你选择今天要通报的那个文件夹
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择你要通报的日期文件夹（例如 CF-20260316车辆违规处罚通知单）";

                // 如果你选好了文件夹并点击了“确定”
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;

                    // 获取这个文件夹里的所有 PDF 附件（只读PDF，发群里也发PDF）
                    string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf");

                    if (pdfFiles.Length == 0)
                    {
                        MessageBox.Show("该文件夹内没有找到 PDF 文件！\n请先使用【Word 导出 PDF】功能生成最终附件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 2. 从文件夹名字里提取出年月日，比如提取出 2026 03 16
                    string folderName = new DirectoryInfo(folderPath).Name;
                    Match dateMatch = Regex.Match(folderName, @"CF-(\d{4})(\d{2})(\d{2})");
                    string displayDate = "某年某月某日";
                    if (dateMatch.Success)
                    {
                        // 拼装成：2026年3月16日 (去掉月份前面多余的 0)
                        displayDate = $"{dateMatch.Groups[1].Value}年{int.Parse(dateMatch.Groups[2].Value)}月{int.Parse(dateMatch.Groups[3].Value)}日";
                    }

                    // 3. 统计车牌和违规次数（揪出多次违规的“惯犯”）
                    Dictionary<string, int> plateCounts = new Dictionary<string, int>();
                    foreach (string file in pdfFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        // 用正则从文件名抓取车牌
                        Match plateMatch = Regex.Match(fileName, @"(粤[A-Z][A-Z0-9]{4,5})");
                        if (plateMatch.Success)
                        {
                            string plate = plateMatch.Groups[1].Value;
                            if (plateCounts.ContainsKey(plate))
                                plateCounts[plate]++; // 违规次数 +1
                            else
                                plateCounts[plate] = 1; // 第一次记为 1
                        }
                    }

                    // 4. 拼装“涉及车辆”这段文字
                    List<string> carList = new List<string>();
                    foreach (var kvp in plateCounts)
                    {
                        // 如果字典里记录的次数大于1，就加上“（多次）”后缀
                        if (kvp.Value > 1)
                            carList.Add($"{kvp.Key}（多次）");
                        else
                            carList.Add(kvp.Key);
                    }
                    string carsString = string.Join("、", carList); // 用顿号连起来

                    // 5. 生成最终的通报文案
                    string reportText = $@"【关于{displayDate}校车路线偏离违规的处理通报】

各位管理员，大家好：

附件是{displayDate}校车运行期间的路线偏离违规处理通知书（共{pdfFiles.Length}份），请及时查收。
涉及车辆：{carsString}。
请各位管理员配合跟进以下工作：

    将违规单下发给对应驾驶员签字确认。

    核实偏离原因（如遇修路、封路等客观原因，请管理员提供证明报备）。

    再次向车队强调：严禁私自更改交警审批路线。

收到请回复，并尽快落实反馈，辛苦大家！";

                    // 6. 神奇的一步：直接把文案塞进电脑剪贴板
                    Clipboard.SetText(reportText);

                    MessageBox.Show($"群通报文案已成功生成！\n包含 {pdfFiles.Length} 份通知单信息。\n\n已自动复制到了你的剪贴板，快去微信群里按 Ctrl+V 粘贴吧！", "生成成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e) // 注意：你的可能叫别的名字，以你双击生成的为准
        {
            // 1. 弹出一个窗口，让你选择当天的日期文件夹
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择包含处罚单的文件夹（程序会自动剔除Word，只打包PDF）";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string sourceFolder = folderDialog.SelectedPath;

                    // 2. 【核心】只把后缀为 .pdf 的文件抓出来！
                    string[] pdfFiles = Directory.GetFiles(sourceFolder, "*.pdf");

                    if (pdfFiles.Length == 0)
                    {
                        MessageBox.Show("该文件夹内没有找到任何 PDF 文件！\n请确认已经点击过【Word转PDF】。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 3. 准备生成压缩包的路径和名字（跟原文件夹同名，直接放在桌面上）
                    string folderName = new DirectoryInfo(sourceFolder).Name;
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string zipFilePath = Path.Combine(desktopPath, $"{folderName}.zip");

                    try
                    {
                        // 如果桌面上已经有一个旧的同名压缩包，先把它删掉，防止文件重复追加
                        if (System.IO.File.Exists(zipFilePath))
                        {
                            System.IO.File.Delete(zipFilePath);
                        }

                        // 4. 创建一个新的 ZIP 压缩包
                        using (ZipArchive archive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
                        {
                            // 遍历刚才找到的所有 PDF，一个个塞进压缩包里
                            foreach (string pdfFile in pdfFiles)
                            {
                                // Path.GetFileName(pdfFile) 确保塞进压缩包的只有文件名，不带长长的电脑路径
                                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
                            }
                        }

                        // 5. 搞定提示
                        MessageBox.Show($"压缩包生成成功！\n\n共打包了 {pdfFiles.Length} 份纯 PDF 文件。\n\n压缩包已直接保存在桌面：\n【{folderName}.zip】\n\n现在可以直接拖拽到微信群里发送了！",
                                        "打包成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("生成压缩包时发生错误，可能是文件正在被别的软件占用：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void btnMatchAdmin_Click(object sender, EventArgs e)
        {
            // 1. 设置台账路径
            string excelPath = Path.Combine(Application.StartupPath, "车牌管理员.xlsx");

            if (!System.IO.File.Exists(excelPath))
            {
                MessageBox.Show("请确保程序目录下有【车牌管理员.xlsx】！", "找不到台账", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. 选择处罚单文件夹
            string folderPath = "";
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择要匹配的处罚单文件夹（对内完整版）";
                if (folderDialog.ShowDialog() != DialogResult.OK) return;
                folderPath = folderDialog.SelectedPath;
            }

            try
            {
                // 3. 提取业务日期（从文件夹名 CF-20260317 中提取）
                string folderName = new DirectoryInfo(folderPath).Name;
                Match dateMatch = Regex.Match(folderName, @"CF-(\d{4})(\d{2})(\d{2})");
                string displayDate = "未知日期";
                if (dateMatch.Success)
                {
                    displayDate = $"{dateMatch.Groups[1].Value}年{int.Parse(dateMatch.Groups[2].Value)}月{int.Parse(dateMatch.Groups[3].Value)}日";
                }

                // 4. 读取 Excel 全量数据
                var vehicleInfoMap = new Dictionary<string, string>();
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        var table = result.Tables[0];

                        for (int i = 1; i < table.Rows.Count; i++)
                        {
                            string plate = table.Rows[i][0]?.ToString().Trim() ?? ""; // A列：车牌
                            string type = table.Rows[i][1]?.ToString().Trim() ?? "未知"; // B列：性质（直接读取，不转换）
                            string admin = table.Rows[i][2]?.ToString().Trim() ?? "未指派"; // C列：管理员
                            string owner = table.Rows[i][3]?.ToString().Trim() ?? "未知"; // D列：车主
                            string unit = table.Rows[i][4]?.ToString().Trim() ?? "未知单位";  // E列：使用单位

                            if (!string.IsNullOrEmpty(plate) && !vehicleInfoMap.ContainsKey(plate))
                            {
                                // 【全量显示】管理员和单位在前，车主和原始性质在后
                                string detail = $"【管理员：{admin} | 单位：{unit} | 车主：{owner} | 性质：{type}】";
                                vehicleInfoMap.Add(plate, detail);
                            }
                        }
                    }
                }

                // 5. 生成清单
                string[] files = Directory.GetFiles(folderPath, "*.pdf");
                System.Text.StringBuilder resultList = new System.Text.StringBuilder();
                resultList.AppendLine($"--- {displayDate} 内部详细归属清单 ---");

                foreach (string file in files)
                {
                    string fileName = Path.GetFileName(file);
                    Match plateMatch = Regex.Match(fileName, @"(粤[A-Z][A-Z0-9]{4,5})");

                    if (plateMatch.Success)
                    {
                        string plate = plateMatch.Groups[1].Value;
                        if (vehicleInfoMap.ContainsKey(plate))
                        {
                            resultList.AppendLine($"车牌：{plate}");
                            resultList.AppendLine($"  => {vehicleInfoMap[plate]}");
                            resultList.AppendLine("--------------------------------------------------");
                        }
                        else
                        {
                            resultList.AppendLine($"车牌：{plate} => [警报] 资料库中未记录此车！");
                        }
                    }
                }

                // 6. 复制并弹窗
                string finalResult = resultList.ToString();
                Clipboard.SetText(finalResult);
                MessageBox.Show(finalResult + "\n\n详细清单已复制到剪贴板！", "内部匹配成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理失败，请检查 Excel 是否被占用：\n" + ex.Message);
            }
        }



        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            // 1. 直接获取程序运行的根目录
            string folderPath = Application.StartupPath;

            try
            {
                // 2. 检查目录是否存在（虽然运行中的程序目录肯定在，但加个保护更稳）
                if (System.IO.Directory.Exists(folderPath))
                {
                    // 3. 使用资源管理器打开该文件夹
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "explorer.exe", // 调用 Windows 资源管理器
                        Arguments = folderPath,    // 传入文件夹路径
                        UseShellExecute = true
                    };
                    System.Diagnostics.Process.Start(psi);
                }
                else
                {
                    MessageBox.Show("找不到程序运行目录！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("无法打开文件夹：\n" + ex.Message, "启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // 调用我们的“验证门卫”
            if (!CheckTimePassword())
            {
                // 密码不对，直接强行关闭程序
                Application.Exit();
            }

        }

        private bool CheckTimePassword()
        {
            // 1. 获取小仓库里存储的到期时间
            DateTime currentExpiry = Properties.Settings.Default.ExpiryDate;

            // 2. 【核心改进】如果还没过期，直接返回 true，不弹窗，实现“静默进入”
            if (DateTime.Now < currentExpiry)
            {
                return true;
            }

            // 3. 如果运行到这里，说明已经过期了，才弹出输入框
            string prompt = $"授权已过期(上次有效期至：{currentExpiry.ToShortDateString()})\n请输入授权码解锁：";
            string userInput = Microsoft.VisualBasic.Interaction.InputBox(prompt, "安全验证", "", -1, -1);

            // --- 权限 A：万能码 admin888 (手动改时间) ---
            if (userInput == "admin888")
            {
                string newDateStr = Microsoft.VisualBasic.Interaction.InputBox("【管理模式】请输入新的到期日期 (yyyy-mm-dd):", "延长授权", DateTime.Now.AddMonths(1).ToString("yyyy-MM-dd"), -1, -1);
                if (DateTime.TryParse(newDateStr, out DateTime newDate))
                {
                    SaveNewExpiry(newDate);
                    return true;
                }
                return false;
            }

            // --- 权限 B：半年特权码 hph666 (直接加半年) ---
            if (userInput == "hph666")
            {
                DateTime halfYearDate = DateTime.Now.AddMonths(6); // 当前时间往后推6个月
                SaveNewExpiry(halfYearDate);
                MessageBox.Show($"特权解锁成功！欢迎回来，Han Peixao。\n软件有效期已延长至：{halfYearDate.ToLongDateString()}", "长期授权成功");
                return true;
            }

            // --- 权限 C：普通月度码 (yyyyMM + hph) ---
            string monthPart = DateTime.Now.ToString("yyyyMM");
            if (userInput == monthPart + "hph")
            {
                // 解锁到本月最后一天
                DateTime endOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
                SaveNewExpiry(endOfMonth);
                return true;
            }

            // 密码错误
            MessageBox.Show("授权码无效，请联系管理员！", "验证失败");
            return false;
        }

        // 提取一个保存时间的小方法，减少重复代码
        private void SaveNewExpiry(DateTime newDate)
        {
            Properties.Settings.Default.ExpiryDate = newDate;
            Properties.Settings.Default.Save();
        }

        private void btnMatchExternal_Click(object sender, EventArgs e)
        {
            // 1. 设置台账路径
            string excelPath = Path.Combine(Application.StartupPath, "车牌管理员.xlsx");

            if (!System.IO.File.Exists(excelPath))
            {
                MessageBox.Show("请确保程序目录下有【车牌管理员.xlsx】！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. 选择处罚单文件夹
            string folderPath = "";
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择要通报的处罚单文件夹";
                if (folderDialog.ShowDialog() != DialogResult.OK) return;
                folderPath = folderDialog.SelectedPath;
            }

            try
            {
                // 3. 提取业务日期
                string folderName = new DirectoryInfo(folderPath).Name;
                Match dateMatch = Regex.Match(folderName, @"CF-(\d{4})(\d{2})(\d{2})");
                string displayDate = "未知日期";
                if (dateMatch.Success)
                {
                    displayDate = $"{dateMatch.Groups[1].Value}年{int.Parse(dateMatch.Groups[2].Value)}月{int.Parse(dateMatch.Groups[3].Value)}日";
                }

                // 4. 读取 Excel 数据（仅提取对外公开的三项）
                var vehicleInfoMap = new Dictionary<string, string>();
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        var table = result.Tables[0];

                        for (int i = 1; i < table.Rows.Count; i++)
                        {
                            string plate = table.Rows[i][0]?.ToString().Trim() ?? ""; // A列：车牌
                            string admin = table.Rows[i][2]?.ToString().Trim() ?? "未指派"; // C列：管理员
                            string unit = table.Rows[i][4]?.ToString().Trim() ?? "未知单位";  // E列：使用单位

                            if (!string.IsNullOrEmpty(plate) && !vehicleInfoMap.ContainsKey(plate))
                            {
                                // 【对外版】只保留三项核心信息
                                string detail = $"【管理员：{admin} | 使用单位：{unit}】";
                                vehicleInfoMap.Add(plate, detail);
                            }
                        }
                    }
                }

                // 5. 生成对外通报清单
                string[] files = Directory.GetFiles(folderPath, "*.pdf");
                System.Text.StringBuilder resultList = new System.Text.StringBuilder();
                resultList.AppendLine($"--- {displayDate} 校车违规运行通报清单 ---");

                foreach (string file in files)
                {
                    string fileName = Path.GetFileName(file);
                    Match plateMatch = Regex.Match(fileName, @"(粤[A-Z][A-Z0-9]{4,5})");

                    if (plateMatch.Success)
                    {
                        string plate = plateMatch.Groups[1].Value;
                        if (vehicleInfoMap.ContainsKey(plate))
                        {
                            resultList.AppendLine($"车牌：{plate} {vehicleInfoMap[plate]}");
                            resultList.AppendLine("--------------------------------------------------");
                        }
                    }
                }

                // 6. 复制并弹窗提示
                string finalResult = resultList.ToString();
                Clipboard.SetText(finalResult);
                MessageBox.Show(finalResult + "\n\n外部通报清单已复制到剪贴板！", "生成成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理失败：\n" + ex.Message);
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            System.Text.StringBuilder helpText = new System.Text.StringBuilder();

            helpText.AppendLine("?? 【全功能按键使用说明词典】\n");

            helpText.AppendLine("?? [读取剪贴板文字]：利用OCR识别剪贴板图片，自动抓取车牌号和违规时间。");
            helpText.AppendLine("?? [导入模板]：初次使用或需更换格式时，选择你的处罚单Word模板。");
            helpText.AppendLine("?? [确认生成处罚单]：将左侧数据与证据图填入模板，生成Word处罚单。");
            helpText.AppendLine("?? [WORD转PDF]：一键批量将当天的Word处罚单转换为正式的PDF格式。");

            helpText.AppendLine("\n--------------------------------------------------");

            helpText.AppendLine("?? [内部详细匹配]：对内管理模式。显示车牌对应的管理员、单位、车主及性质(挂靠/自己)。");
            helpText.AppendLine("?? [生成外部通报清单]：对外通报模式。仅保留车牌、管理员、单位，保护内部隐私。");
            helpText.AppendLine("?? [生成群通报]：自动统计全天数据，生成一段发在微信大群的简易通报文案。");
            helpText.AppendLine("?? [生成压缩包]：智能筛选文件夹，只把正式的PDF打包成ZIP，方便群发。");

            helpText.AppendLine("\n--------------------------------------------------");

            helpText.AppendLine("?? [打开本地台账]：打开程序根目录。可修改Excel名单、查看模板、清理过期文件。");
            helpText.AppendLine("?? [添加快捷方式]：在桌面生成软件图标，无需进入开发目录寻找exe。");

            helpText.AppendLine("\n?? 小技巧：输入半年特权码 hph666 (直接加半年) 可修改有效期。");

            MessageBox.Show(helpText.ToString(), "Fine-Ticket 功能词典", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSingleMatch_Click(object sender, EventArgs e)
        {
            // 1. 获取你要搜索的车牌后四位 (这里换成了你真实的文本框名字 txtPlate)
            string searchCarId = txtPlate.Text.Trim();
            if (string.IsNullOrEmpty(searchCarId))
            {
                MessageBox.Show("请先在左侧输入车牌号（如 0811）！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 2. 直接使用你现有的台账路径逻辑
            string excelPath = Path.Combine(Application.StartupPath, "车牌管理员.xlsx");

            if (!System.IO.File.Exists(excelPath))
            {
                MessageBox.Show("请确保程序目录下有【车牌管理员.xlsx】！", "找不到台账", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        var table = result.Tables[0];
                        bool isFound = false;

                        // 3. 遍历台账寻找这辆车
                        for (int i = 1; i < table.Rows.Count; i++)
                        {
                            // A列：车牌
                            string plate = table.Rows[i][0]?.ToString().Trim() ?? "";

                            // 只要台账里的车牌包含了你输入的数字（比如 0811）
                            if (!string.IsNullOrEmpty(plate) && plate.Contains(searchCarId))
                            {
                                // 提取对应的其余信息
                                string type = table.Rows[i][1]?.ToString().Trim() ?? "未知";
                                string admin = table.Rows[i][2]?.ToString().Trim() ?? "未指派";
                                string owner = table.Rows[i][3]?.ToString().Trim() ?? "未知";
                                string unit = table.Rows[i][4]?.ToString().Trim() ?? "未知单位";

                                // 4. 找到后直接弹窗展示详细信息
                                string detailMsg = $"✅ 匹配成功！\n\n" +
                                                   $"【车牌】：{plate}\n" +
                                                   $"【单位】：{unit}\n" +
                                                   $"【管理员】：{admin}\n" +
                                                   $"【车主】：{owner}\n" +
                                                   $"【性质】：{type}";

                                MessageBox.Show(detailMsg, "单车匹配结果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                isFound = true;
                                break; // 找到了就立刻停止扫描，秒出结果
                            }
                        }

                        if (!isFound)
                        {
                            MessageBox.Show($"在台账中未找到包含【{searchCarId}】的车辆记录！", "查无此车", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取失败，请检查 Excel 是否正被打开占用：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBatchOcrMatch_Click(object sender, EventArgs e)
        {
            // 1. 检查剪贴板中是否有文本
            if (!Clipboard.ContainsText())
            {
                MessageBox.Show("剪贴板中没有文字！请先用微信截图整个报警列表，提取并复制文字。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string clipboardText = Clipboard.GetText();

            // 2. 正则匹配提取所有车牌，并统计报警次数
            // 注意这里用的是 Matches (复数)，抓取所有的车牌
            MatchCollection plateMatches = Regex.Matches(clipboardText, @"(粤[A-Z][A-Z0-9]{4,5})");

            if (plateMatches.Count == 0)
            {
                MessageBox.Show("未能从剪贴板文本中识别到任何车牌，请确认截图包含了车牌列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 用字典统计每个车牌出现的次数 (车牌号 -> 报警次数)
            Dictionary<string, int> violationCounts = new Dictionary<string, int>();
            foreach (Match m in plateMatches)
            {
                string plate = m.Groups[1].Value;
                if (violationCounts.ContainsKey(plate))
                    violationCounts[plate]++;
                else
                    violationCounts[plate] = 1;
            }

            // 3. 读取本地 Excel 台账获取管理员和学校信息
            string excelPath = Path.Combine(Application.StartupPath, "车牌管理员.xlsx");
            if (!System.IO.File.Exists(excelPath))
            {
                MessageBox.Show("请确保程序目录下有【车牌管理员.xlsx】！", "找不到台账", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 字典：车牌号 -> string[] { 管理员, 学校单位 }
            Dictionary<string, string[]> vehicleInfoMap = new Dictionary<string, string[]>();

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        var table = result.Tables[0];

                        for (int i = 1; i < table.Rows.Count; i++) // 跳过表头
                        {
                            string plate = table.Rows[i][0]?.ToString().Trim() ?? "";   // A列: 车牌
                            string admin = table.Rows[i][2]?.ToString().Trim() ?? "未指派"; // C列: 管理员
                            string unit = table.Rows[i][4]?.ToString().Trim() ?? "未知单位";  // E列: 使用单位

                            if (!string.IsNullOrEmpty(plate) && !vehicleInfoMap.ContainsKey(plate))
                            {
                                vehicleInfoMap.Add(plate, new string[] { admin, unit });
                            }
                        }
                    }
                }

                // 4. 构建支持 Excel 直接粘贴的字符串 (利用 \t 制表符自动分列)
                System.Text.StringBuilder excelData = new System.Text.StringBuilder();

                // 构建你期望的表头
                excelData.AppendLine("序号\t车牌\t管理员\t使用学校\t路段偏离报警(次)");

                int index = 1;
                foreach (var kvp in violationCounts)
                {
                    string plate = kvp.Key;
                    int count = kvp.Value;
                    string admin = "未匹配到";
                    string unit = "未匹配到";

                    // 如果台账里有这个车牌，就取出来
                    if (vehicleInfoMap.ContainsKey(plate))
                    {
                        admin = vehicleInfoMap[plate][0];
                        unit = vehicleInfoMap[plate][1];
                    }

                    // 用 \t 隔开每一列
                    excelData.AppendLine($"{index}\t{plate}\t{admin}\t{unit}\t{count}");
                    index++;
                }

                // 5. 将拼好的数据塞入剪贴板
                Clipboard.SetText(excelData.ToString());

                MessageBox.Show($"成功识别到 {violationCounts.Count} 台违规车辆！\n\n数据已自动排版并复制到了剪贴板。\n现在你可以直接新建一个 Excel 表格，选中 A1 单元格，按下 Ctrl+V 直接粘贴了！", "批量处理成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取台账时发生错误，请检查 Excel 是否正被打开占用：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
