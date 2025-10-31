using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Text.RegularExpressions;

namespace LabelGenerator
{
    public partial class MainForm : Form
    {
        private TextBox txtItemName;
        private TextBox txtQuantity;
        private TextBox txtStartNumber;
        private TextBox txtLocation;
        private Button btnGenerate;
        private Button btnAddToBatch;
        private DataGridView dgvBatchItems;
        private Button btnGenerateBatch;
        private Button btnClearBatch;
        private Label lblStatus;
        private TabControl tabControl;

        private Button btnImportCsv; // 新增CSV导入按钮
        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "物品标签生成器";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 创建 TabControl
            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            this.Controls.Add(tabControl);

            // 单个物品标签页
            TabPage tabSingle = new TabPage("单个物品");
            tabControl.TabPages.Add(tabSingle);
            CreateSingleItemTab(tabSingle);

            // 批量物品标签页
            TabPage tabBatch = new TabPage("批量物品");
            tabControl.TabPages.Add(tabBatch);
            CreateBatchItemTab(tabBatch);

            // 状态栏
            lblStatus = new Label();
            lblStatus.Dock = DockStyle.Bottom;
            lblStatus.Height = 30;
            lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblStatus.Text = "就绪";
            this.Controls.Add(lblStatus);
        }

        private void CreateSingleItemTab(TabPage tab)
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(20);
            tab.Controls.Add(panel);

            int yPos = 20;
            int labelWidth = 120;
            int textBoxWidth = 300;
            int spacing = 40;

            // 物品名字
            Label lblName = new Label();
            lblName.Text = "物品名字:";
            lblName.Location = new System.Drawing.Point(20, yPos);
            lblName.Size = new System.Drawing.Size(labelWidth, 25);
            panel.Controls.Add(lblName);

            txtItemName = new TextBox();
            txtItemName.Location = new System.Drawing.Point(150, yPos);
            txtItemName.Size = new System.Drawing.Size(textBoxWidth, 25);
            panel.Controls.Add(txtItemName);

            yPos += spacing;

            // 物品数量
            Label lblQuantity = new Label();
            lblQuantity.Text = "物品数量:";
            lblQuantity.Location = new System.Drawing.Point(20, yPos);
            lblQuantity.Size = new System.Drawing.Size(labelWidth, 25);
            panel.Controls.Add(lblQuantity);

            txtQuantity = new TextBox();
            txtQuantity.Location = new System.Drawing.Point(150, yPos);
            txtQuantity.Size = new System.Drawing.Size(textBoxWidth, 25);
            panel.Controls.Add(txtQuantity);

            yPos += spacing;

            // 起始编号
            Label lblStartNumber = new Label();
            lblStartNumber.Text = "起始编号:";
            lblStartNumber.Location = new System.Drawing.Point(20, yPos);
            lblStartNumber.Size = new System.Drawing.Size(labelWidth, 25);
            panel.Controls.Add(lblStartNumber);

            txtStartNumber = new TextBox();
            txtStartNumber.Location = new System.Drawing.Point(150, yPos);
            txtStartNumber.Size = new System.Drawing.Size(textBoxWidth, 25);
            panel.Controls.Add(txtStartNumber);

            yPos += spacing;

            // 所在室
            Label lblLocation = new Label();
            lblLocation.Text = "所在室:";
            lblLocation.Location = new System.Drawing.Point(20, yPos);
            lblLocation.Size = new System.Drawing.Size(labelWidth, 25);
            panel.Controls.Add(lblLocation);

            txtLocation = new TextBox();
            txtLocation.Location = new System.Drawing.Point(150, yPos);
            txtLocation.Size = new System.Drawing.Size(textBoxWidth, 25);
            panel.Controls.Add(txtLocation);

            yPos += spacing + 20;

            // 生成按钮
            btnGenerate = new Button();
            btnGenerate.Text = "生成标签";
            btnGenerate.Location = new System.Drawing.Point(150, yPos);
            btnGenerate.Size = new System.Drawing.Size(150, 35);
            btnGenerate.Click += BtnGenerate_Click;
            panel.Controls.Add(btnGenerate);
        }

        private void CreateBatchItemTab(TabPage tab)
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(20);
            tab.Controls.Add(panel);

            // 说明标签
            Label lblInfo = new Label();
            lblInfo.Text = "在下方表格中输入多个物品信息，或从CSV文件导入";
            lblInfo.Location = new System.Drawing.Point(20, 20);
            lblInfo.Size = new System.Drawing.Size(500, 25);
            panel.Controls.Add(lblInfo);

            // DataGridView
            dgvBatchItems = new DataGridView();
            dgvBatchItems.Location = new System.Drawing.Point(20, 50);
            dgvBatchItems.Size = new System.Drawing.Size(740, 350);
            dgvBatchItems.AllowUserToAddRows = true;
            dgvBatchItems.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dgvBatchItems.Columns.Add("ItemName", "物品名字");
            dgvBatchItems.Columns.Add("Quantity", "数量");
            dgvBatchItems.Columns.Add("StartNumber", "起始编号");
            dgvBatchItems.Columns.Add("Location", "所在室");

            dgvBatchItems.Columns[0].Width = 200;
            dgvBatchItems.Columns[1].Width = 100;
            dgvBatchItems.Columns[2].Width = 150;
            dgvBatchItems.Columns[3].Width = 200;

            panel.Controls.Add(dgvBatchItems);

            // 按钮面板
            Panel buttonPanel = new Panel();
            buttonPanel.Location = new System.Drawing.Point(20, 410);
            buttonPanel.Size = new System.Drawing.Size(740, 50);
            panel.Controls.Add(buttonPanel);

            btnGenerateBatch = new Button();
            btnGenerateBatch.Text = "生成批量标签";
            btnGenerateBatch.Location = new System.Drawing.Point(0, 0);
            btnGenerateBatch.Size = new System.Drawing.Size(150, 35);
            btnGenerateBatch.Click += BtnGenerateBatch_Click;
            buttonPanel.Controls.Add(btnGenerateBatch);

            btnClearBatch = new Button();
            btnClearBatch.Text = "清空表格";
            btnClearBatch.Location = new System.Drawing.Point(160, 0);
            btnClearBatch.Size = new System.Drawing.Size(150, 35);
            btnClearBatch.Click += (s, e) => dgvBatchItems.Rows.Clear();
            buttonPanel.Controls.Add(btnClearBatch);

            // 新增CSV导入按钮
            btnImportCsv = new Button();
            btnImportCsv.Text = "从CSV导入";
            btnImportCsv.Location = new System.Drawing.Point(320, 0); // 放在清空按钮旁边
            btnImportCsv.Size = new System.Drawing.Size(150, 35);
            btnImportCsv.Click += BtnImportCsv_Click;
            buttonPanel.Controls.Add(btnImportCsv);
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                // 验证输入
                if (string.IsNullOrWhiteSpace(txtItemName.Text))
                {
                    MessageBox.Show("请输入物品名字", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!int.TryParse(txtQuantity.Text, out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("请输入有效的数量（正整数）", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. 获取起始编号的 *原始字符串*
                string startNumberStr = txtStartNumber.Text.Trim();

                if (!int.TryParse(startNumberStr, out int startNumber)) // 🆕 验证原始字符串
                {
                    MessageBox.Show("请输入有效的起始编号（整数）", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (string.IsNullOrWhiteSpace(txtLocation.Text))
                {
                    MessageBox.Show("请输入所在室", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 2. 检查是否需要补零
                // (条件：以"0"开头，且总长度大于1，例如 "01", "001")
                bool usePadding = startNumberStr.StartsWith("0") && startNumberStr.Length > 1;
                int paddingLength = startNumberStr.Length; // 原始长度，例如 "001" -> 3
                string formatString = "D" + paddingLength; // 生成格式化字符串, 例如 "D3"

                // 创建物品列表
                List<ItemInfo> items = new List<ItemInfo>();
                for (int i = 0; i < quantity; i++)
                {
                    int currentNumber = startNumber + i;

                    // 3. 根据需要格式化编号
                    string numberString;
                    if (usePadding)
                    {
                        // .ToString("D3") 会将 1 变为 "001", 10 变为 "010", 100 变为 "100"
                        numberString = currentNumber.ToString(formatString);
                    }
                    else
                    {
                        // 保持原来的逻辑
                        numberString = currentNumber.ToString();
                    }

                    items.Add(new ItemInfo
                    {
                        Name = txtItemName.Text,
                        Number = numberString, // 4. 使用格式化后的编号
                        Location = txtLocation.Text
                    });
                }

                // 生成标签
                GenerateLabels(items);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "生成失败";
            }
        }

        private void BtnGenerateBatch_Click(object sender, EventArgs e)
        {
            try
            {
                List<ItemInfo> items = new List<ItemInfo>();

                foreach (DataGridViewRow row in dgvBatchItems.Rows)
                {
                    if (row.IsNewRow) continue;

                    string name = row.Cells[0].Value?.ToString();
                    string quantityStr = row.Cells[1].Value?.ToString();

                    // 🆕 1. 获取起始编号的 *原始字符串* (并 Trim)
                    string startNumberStr = row.Cells[2].Value?.ToString()?.Trim();

                    string location = row.Cells[3].Value?.ToString();

                    if (string.IsNullOrWhiteSpace(name) ||
                        string.IsNullOrWhiteSpace(quantityStr) ||
                        string.IsNullOrWhiteSpace(startNumberStr) || // 🆕 验证原始字符串
                        string.IsNullOrWhiteSpace(location))
                    {
                        continue;
                    }

                    if (!int.TryParse(quantityStr, out int quantity) || quantity <= 0)
                    {
                        MessageBox.Show($"行 {row.Index + 1}: 数量无效", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (!int.TryParse(startNumberStr, out int startNumber)) // 🆕 验证原始字符串
                    {
                        MessageBox.Show($"行 {row.Index + 1}: 起始编号无效", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 🆕 2. 检查是否需要补零 (逻辑同上)
                    bool usePadding = startNumberStr.StartsWith("0") && startNumberStr.Length > 1;
                    int paddingLength = startNumberStr.Length;
                    string formatString = "D" + paddingLength;

                    for (int i = 0; i < quantity; i++)
                    {
                        int currentNumber = startNumber + i;

                        // 🆕 3. 根据需要格式化编号
                        string numberString = usePadding
                            ? currentNumber.ToString(formatString)
                            : currentNumber.ToString();

                        items.Add(new ItemInfo
                        {
                            Name = name,
                            Number = numberString, // 🆕 4. 使用格式化后的编号
                            Location = location
                        });
                    }
                }

                if (items.Count == 0)
                {
                    MessageBox.Show("请至少输入一行有效数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                GenerateLabels(items);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "生成失败";
            }
        }

        private void GenerateLabels(List<ItemInfo> items)
        {
            lblStatus.Text = "正在生成标签...";
            Application.DoEvents();

            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "template.docx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("未找到 template.docx 文件，请确保模板文件在程序目录下", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "模板文件不存在";
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputDocxPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"标签_{timestamp}.docx");
            string outputPdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"标签_{timestamp}.pdf");
            string outputJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"标签_{timestamp}.json");
            string outputCsvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"标签_{timestamp}.csv");

            // 创建新文档
            using (DocX firstTemplate = DocX.Load(templatePath))
            {
                firstTemplate.ReplaceText("物品名字", items[0].Name);
                firstTemplate.ReplaceText("物品编号", items[0].Number);
                firstTemplate.ReplaceText("所在", items[0].Location);

                // 2. 【保存为最终文件】: 不是 Create，而是 SaveAs！
                // 这一步决定了 outputDocxPath 具有模板的正确页面设置
                firstTemplate.SaveAs(outputDocxPath);
            }


            if (items.Count > 1)
            {
                // 4. 【加载刚刚保存的最终文件】
                using (DocX finalDocument = DocX.Load(outputDocxPath))
                {
                    // 5. 【循环剩余物品】: 注意循环是从 i = 1 (第二个) 开始
                    for (int i = 1; i < items.Count; i++)
                    {
                        // 为每个新物品 *重新加载* 一份干净的模板
                        using (DocX subsequentTemplate = DocX.Load(templatePath))
                        {
                            subsequentTemplate.ReplaceText("物品名字", items[i].Name);
                            subsequentTemplate.ReplaceText("物品编号", items[i].Number);
                            subsequentTemplate.ReplaceText("所在", items[i].Location);

                            // 6. 【追加文档】
                            finalDocument.InsertDocument(subsequentTemplate);
                        }
                    }

                    // 7. 【保存所有追加】
                    finalDocument.Save();
                }
            }
            Application.DoEvents();

            // 转换为 PDF
            //ConvertDocxToPdf(outputDocxPath, outputPdfPath);

            lblStatus.Text = "正在生成JSON和CSV...";
            Application.DoEvents();

            // 生成 JSON
            string json = JsonSerializer.Serialize(items, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
            File.WriteAllText(outputJsonPath, json, Encoding.UTF8);

            // 生成 CSV
            StringBuilder csv = new StringBuilder();
            csv.AppendLine("\"物品名字\",\"物品编号\",\"所在室\""); // 确保表头也带引号
            foreach (var item in items)
            {
                // 使用 " " 来包裹，确保CSV格式的健壮性
                csv.AppendLine($"\"{item.Name}\",\"{item.Number}\",\"{item.Location}\"");
            }
            File.WriteAllText(outputCsvPath, csv.ToString(), Encoding.UTF8);

            lblStatus.Text = "生成完成";
            MessageBox.Show($"标签生成成功！\n\n生成文件:\n{Path.GetFileName(outputDocxPath)}\n{Path.GetFileName(outputPdfPath)}\n{Path.GetFileName(outputJsonPath)}\n{Path.GetFileName(outputCsvPath)}",
                "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        // CSV导入按钮的点击事件
        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "CSV 文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                    openFileDialog.Title = "选择要导入的CSV文件";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;
                        lblStatus.Text = "正在从CSV导入...";
                        Application.DoEvents();

                        List<ItemInfo> items = new List<ItemInfo>();

                        // 读取所有行, 跳过表头 (Skip(1))
                        var lines = File.ReadAllLines(filePath, Encoding.UTF8).Skip(1);
                        int lineCount = 1; // 用于错误提示

                        foreach (string line in lines)
                        {
                            lineCount++;
                            if (string.IsNullOrWhiteSpace(line)) continue;

                            // 你的CSV输出格式是 "Name","Number","Location"
                            // 我们用一个简单的方法来解析这种带引号的格式
                            // 正则表达式匹配被引号包裹的内容
                            MatchCollection matches = Regex.Matches(line, "\"(.*?)\"");

                            if (matches.Count >= 3)
                            {
                                string name = matches[0].Groups[1].Value;
                                string number = matches[1].Groups[1].Value;
                                string location = matches[2].Groups[1].Value;

                                items.Add(new ItemInfo
                                {
                                    Name = name,
                                    Number = number,
                                    Location = location
                                });
                            }
                            else
                            {
                                // 如果正则不匹配，尝试简单的逗号分割（不带引号的CSV）
                                string[] parts = line.Split(',');
                                if (parts.Length >= 3)
                                {
                                    items.Add(new ItemInfo
                                    {
                                        Name = parts[0].Trim(),
                                        Number = parts[1].Trim(),
                                        Location = parts[2].Trim()
                                    });
                                }
                                else
                                {
                                    lblStatus.Text = $"CSV 第 {lineCount} 行格式错误，已跳过";
                                }
                            }
                        }

                        if (items.Count == 0)
                        {
                            MessageBox.Show("CSV文件为空或格式不正确。\n\n请确保CSV格式为:\n\"物品名字\",\"物品编号\",\"所在室\"\n(带引号，UTF-8编码)", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            lblStatus.Text = "导入失败";
                            return;
                        }

                        // 如果成功加载了物品，直接去生成
                        // 确保你已经应用了 GenerateLabels 的修复
                        GenerateLabels(items);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"CSV导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "导入失败";
            }
        }
    }

    


    public class ItemInfo
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Location { get; set; }
    }
}