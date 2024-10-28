using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace XlsxToLuaGUI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            tbProgramPath.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbExcelFolderPath.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbExportLuaFolderPath.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbClientFolderPath.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbLangFilePath.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbPartExcelNames.DragEnter += new DragEventHandler(_TextBoxDragEnter);
            tbExceptExcelNames.DragEnter += new DragEventHandler(_TextBoxDragEnter);

            tbProgramPath.DragDrop += new DragEventHandler(_TextBoxOneFileDragDrop);
            tbLangFilePath.DragDrop += new DragEventHandler(_TextBoxOneFileDragDrop);
            tbExcelFolderPath.DragDrop += new DragEventHandler(_TextBoxOneDirDragDrop);
            tbExportLuaFolderPath.DragDrop += new DragEventHandler(_TextBoxOneDirDragDrop);
            tbClientFolderPath.DragDrop += new DragEventHandler(_TextBoxOneDirDragDrop);
            tbPartExcelNames.DragDrop += new DragEventHandler(_TextBoxMultipleExcelFileDragDrop);
            tbExceptExcelNames.DragDrop += new DragEventHandler(_TextBoxMultipleExcelFileDragDrop);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // 部分文本框填入默认值
            tbExportLuaFolderPath.Text = AppValues.NOT_EXPORT_LUA_PARAM_STRING;
            tbClientFolderPath.Text = AppValues.NO_CLIENT_PATH_PARAM_STRING;
            tbLangFilePath.Text = AppValues.NO_LANG_PARAM_STRING;
            // 查找本程序所在目录下是否含有XlsxToLua工具，如果有直接填写路径到“工具所在目录”文本框中
            string defaultPath = Utils.CombinePath(AppValues.PROGRAM_FOLDER_PATH, AppValues.PROGRAM_NAME);
            if (File.Exists(defaultPath))
                tbProgramPath.Text = defaultPath;

            string defaultConfig = Utils.CombinePath(AppValues.PROGRAM_FOLDER_PATH, AppValues.CONFIG_FILE_NAME);
            if (File.Exists(defaultConfig))
            {
                string errorString = null;
                Dictionary<string, string> config = TxtConfigReader.ParseTxtConfigFile(defaultConfig, AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR, out errorString);
                if (string.IsNullOrEmpty(errorString))
                {
                    //if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_PROGRAM_PATH))
                    //    tbProgramPath.Text = config[AppValues.SAVE_CONFIG_KEY_PROGRAM_PATH];
                    if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_EXCEL_FOLDER_PATH))
                        tbExcelFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_EXCEL_FOLDER_PATH];
                    if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_EXPORT_LUA_FOLDER_PATH))
                        tbExportLuaFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_EXPORT_LUA_FOLDER_PATH];
                    if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_CLIENT_FOLDER_PATH))
                        tbClientFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_CLIENT_FOLDER_PATH];
                    if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_LANG_FILE_PATH))
                        tbLangFilePath.Text = config[AppValues.SAVE_CONFIG_KEY_LANG_FILE_PATH];
                    if (config.ContainsKey(AppValues.EXPORT_INCLUDE_SUBFOLDER_PARAM_STRING))
                        cbExportIncludeSubfolder.Checked = true;
                    if (config.ContainsKey(AppValues.EXPORT_KEEP_DIRECTORY_STRUCTURE_PARAM_STRING))
                        cbExportKeepDirectoryStructure.Checked = true;
                    if (config.ContainsKey(AppValues.NEED_COLUMN_INFO_PARAM_STRING))
                        cbColumnInfo.Checked = true;
                    if (config.ContainsKey(AppValues.UNCHECKED_PARAM_STRING))
                        cbUnchecked.Checked = true;
                    if (config.ContainsKey(AppValues.LANG_NOT_MATCHING_PRINT_PARAM_STRING))
                        cbPrintEmptyStringWhenLangNotMatching.Checked = true;
                    if (config.ContainsKey(AppValues.EXPORT_MYSQL_PARAM_STRING))
                        cbExportMySQL.Checked = true;
                    if (config.ContainsKey(AppValues.ALLOWED_NULL_NUMBER_PARAM_STRING))
                        cbAllowedNullNumber.Checked = true;
                    if (config.ContainsKey(AppValues.PART_EXPORT_PARAM_STRING))
                        tbPartExcelNames.Text = config[AppValues.PART_EXPORT_PARAM_STRING];
                    if (config.ContainsKey(AppValues.EXCEPT_EXPORT_PARAM_STRING))
                        tbExceptExcelNames.Text = config[AppValues.EXCEPT_EXPORT_PARAM_STRING];

                    cbPart.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_PART);
                    cbExcept.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_EXCEPT);
                    cbIsUseRelativePath.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_USE_RELATIVE_PATH);
                }
            }
        }

        private void btnChooseExcelFolderPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择Excel文件所在目录";
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = dialog.SelectedPath;
                tbExcelFolderPath.Text = folderPath;
            }
        }

        private void btnChooseExportLuaFolderPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择lua文件导出目录";
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = dialog.SelectedPath;
                tbExportLuaFolderPath.Text = folderPath;
            }
        }

        private void btnChooseClientFolderPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择项目Client所在目录";
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = dialog.SelectedPath;
                tbClientFolderPath.Text = folderPath;
            }
        }

        private void btnChooseLangFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "请选择lang文件所在路径";
            dialog.Multiselect = false;
            dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = dialog.FileName;
                tbLangFilePath.Text = filePath;
            }
        }

        private void btnChoosePartExcel_Click(object sender, EventArgs e)
        {
            string errorString = null;
            string chooseExcelNames = _GetChoosePartExcelFile(tbPartExcelNames.Text, out errorString);
            if (errorString == null)
            {
                if (chooseExcelNames != null)
                    tbPartExcelNames.Text = chooseExcelNames;
            }
            else
                MessageBox.Show(errorString, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnChooseExceptExcel_Click(object sender, EventArgs e)
        {
            string errorString = null;
            string exceptExcelNames = _GetChoosePartExcelFile(tbExceptExcelNames.Text, out errorString);
            if (errorString == null)
            {
                if (exceptExcelNames != null)
                    tbExceptExcelNames.Text = exceptExcelNames;
            }
            else
                MessageBox.Show(errorString, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnChooseProgramPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "请选择XlsxToLua工具所在路径";
            dialog.Multiselect = false;
            dialog.Filter = "Exe files (*.exe)|*.exe";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = dialog.FileName;
                tbProgramPath.Text = filePath;
            }
        }

        private void cbUnchecked_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            _WarnWhenChooseDangerousParam(cb);
        }

        private void cbAllowedNullNumber_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            _WarnWhenChooseDangerousParam(cb);
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            if (_CheckConfig() == true)
            {
                string programPath = tbProgramPath.Text.Trim();
                string batContent = _GetExecuteParamString();
                // System.Diagnostics.Process.Start函数无法识别用/分层的相对路径，故需进行转换
                System.Diagnostics.Process.Start(programPath.Replace('/', '\\'), batContent.Substring(batContent.IndexOf(programPath) + programPath.Length + 1));
            }
        }

        private void btnGenerateBat_Click(object sender, EventArgs e)
        {
            if (_CheckConfig() == true)
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Title = "请选择要生成的bat批处理脚本所在路径";
                dialog.InitialDirectory = Path.GetDirectoryName(tbProgramPath.Text.Trim());
                dialog.Filter = "Bat files (*.bat)|*.bat";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = dialog.FileName;
                    string batContent = _GetExecuteParamString();
                    string errorString = null;
                    Utils.SaveFile(filePath, batContent, out errorString);
                    if (string.IsNullOrEmpty(errorString))
                        MessageBox.Show("生成bat批处理脚本成功", "恭喜", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show(string.Format("保存bat批处理脚本失败：{0}", errorString), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            if (_CheckConfig() == true)
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Title = "请选择配置文件的保存路径";
                dialog.InitialDirectory = AppValues.PROGRAM_FOLDER_PATH;
                dialog.Filter = "Config files (*.txt)|*.txt";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    StringBuilder configStringBuilder = new StringBuilder();
                    Uri programUri = new Uri(AppValues.PROGRAM_PATH);

                    string programPath = tbProgramPath.Text.Trim();
                    string excelFolderPath = tbExcelFolderPath.Text.Trim();
                    string exportLuaFolderPath = tbExportLuaFolderPath.Text.Trim();
                    string clientFolderPath = tbClientFolderPath.Text.Trim();
                    string langFilePath = tbLangFilePath.Text.Trim();
                    if (cbIsUseRelativePath.Checked == true)
                    {
                        if (Path.IsPathRooted(programPath))
                            programPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(programPath)).ToString());
                        if (Path.IsPathRooted(excelFolderPath))
                            excelFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(excelFolderPath)).ToString());
                        if (!exportLuaFolderPath.Equals(AppValues.NOT_EXPORT_LUA_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(exportLuaFolderPath))
                            exportLuaFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(exportLuaFolderPath)).ToString());
                        if (!clientFolderPath.Equals(AppValues.NO_CLIENT_PATH_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(clientFolderPath))
                            clientFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(clientFolderPath)).ToString());

                        if (!langFilePath.Equals(AppValues.NO_LANG_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(langFilePath))
                            langFilePath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(langFilePath)).ToString());
                    }
                    configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_PROGRAM_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(programPath);
                    configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_EXCEL_FOLDER_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(excelFolderPath);
                    configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_EXPORT_LUA_FOLDER_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(exportLuaFolderPath);
                    configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_CLIENT_FOLDER_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(clientFolderPath);
                    configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_LANG_FILE_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(langFilePath);

                    string partExcelNames = tbPartExcelNames.Text.Trim();
                    if (!string.IsNullOrEmpty(partExcelNames))
                        configStringBuilder.Append(AppValues.PART_EXPORT_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(partExcelNames);

                    string exceptExcelNames = tbExceptExcelNames.Text.Trim();
                    if (!string.IsNullOrEmpty(exceptExcelNames))
                        configStringBuilder.Append(AppValues.EXCEPT_EXPORT_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(exceptExcelNames);

                    const string TRUE_STRING = "true";
                    if (cbExportIncludeSubfolder.Checked == true)
                        configStringBuilder.Append(AppValues.EXPORT_INCLUDE_SUBFOLDER_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbExportKeepDirectoryStructure.Checked == true)
                        configStringBuilder.Append(AppValues.EXPORT_KEEP_DIRECTORY_STRUCTURE_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbColumnInfo.Checked == true)
                        configStringBuilder.Append(AppValues.NEED_COLUMN_INFO_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbUnchecked.Checked == true)
                        configStringBuilder.Append(AppValues.UNCHECKED_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbPrintEmptyStringWhenLangNotMatching.Checked == true)
                        configStringBuilder.Append(AppValues.LANG_NOT_MATCHING_PRINT_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbExportMySQL.Checked == true)
                        configStringBuilder.Append(AppValues.EXPORT_MYSQL_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbAllowedNullNumber.Checked == true)
                        configStringBuilder.Append(AppValues.ALLOWED_NULL_NUMBER_PARAM_STRING).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);

                    if (cbPart.Checked == true)
                        configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_PART).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbExcept.Checked == true)
                        configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_EXCEPT).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);
                    if (cbIsUseRelativePath.Checked == true)
                        configStringBuilder.Append(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_USE_RELATIVE_PATH).Append(AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR).AppendLine(TRUE_STRING);

                    string errorString = null;
                    Utils.SaveFile(dialog.FileName, configStringBuilder.ToString(), out errorString);
                    if (string.IsNullOrEmpty(errorString))
                        MessageBox.Show("保存配置文件成功", "恭喜", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show(string.Format("保存配置文件失败：{0}", errorString), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnLoadConfig_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "请选择配置文件所在路径";
            dialog.InitialDirectory = AppValues.PROGRAM_FOLDER_PATH;
            dialog.Multiselect = false;
            dialog.Filter = "Config files (*.txt)|*.txt";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string errorString = null;
                Dictionary<string, string> config = TxtConfigReader.ParseTxtConfigFile(dialog.FileName, AppValues.SAVE_CONFIG_KEY_VALUE_SEPARATOR, out errorString);
                if (!string.IsNullOrEmpty(errorString))
                {
                    MessageBox.Show(string.Format("打开配置文件失败：\n{0}", errorString), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_PROGRAM_PATH))
                    tbProgramPath.Text = config[AppValues.SAVE_CONFIG_KEY_PROGRAM_PATH];
                if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_EXCEL_FOLDER_PATH))
                    tbExcelFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_EXCEL_FOLDER_PATH];
                if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_EXPORT_LUA_FOLDER_PATH))
                    tbExportLuaFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_EXPORT_LUA_FOLDER_PATH];
                if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_CLIENT_FOLDER_PATH))
                    tbClientFolderPath.Text = config[AppValues.SAVE_CONFIG_KEY_CLIENT_FOLDER_PATH];
                if (config.ContainsKey(AppValues.SAVE_CONFIG_KEY_LANG_FILE_PATH))
                    tbLangFilePath.Text = config[AppValues.SAVE_CONFIG_KEY_LANG_FILE_PATH];
                if (config.ContainsKey(AppValues.EXPORT_INCLUDE_SUBFOLDER_PARAM_STRING))
                    cbExportIncludeSubfolder.Checked = true;
                if (config.ContainsKey(AppValues.EXPORT_KEEP_DIRECTORY_STRUCTURE_PARAM_STRING))
                    cbExportKeepDirectoryStructure.Checked = true;
                if (config.ContainsKey(AppValues.NEED_COLUMN_INFO_PARAM_STRING))
                    cbColumnInfo.Checked = true;
                if (config.ContainsKey(AppValues.UNCHECKED_PARAM_STRING))
                    cbUnchecked.Checked = true;
                if (config.ContainsKey(AppValues.LANG_NOT_MATCHING_PRINT_PARAM_STRING))
                    cbPrintEmptyStringWhenLangNotMatching.Checked = true;
                if (config.ContainsKey(AppValues.EXPORT_MYSQL_PARAM_STRING))
                    cbExportMySQL.Checked = true;
                if (config.ContainsKey(AppValues.ALLOWED_NULL_NUMBER_PARAM_STRING))
                    cbAllowedNullNumber.Checked = true;
                if (config.ContainsKey(AppValues.PART_EXPORT_PARAM_STRING))
                    tbPartExcelNames.Text = config[AppValues.PART_EXPORT_PARAM_STRING];
                if (config.ContainsKey(AppValues.EXCEPT_EXPORT_PARAM_STRING))
                    tbExceptExcelNames.Text = config[AppValues.EXCEPT_EXPORT_PARAM_STRING];
               
                cbPart.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_PART);
                cbExcept.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_EXCEPT);
                cbIsUseRelativePath.Checked = config.ContainsKey(AppValues.SAVE_CONFIG_KEY_IS_CHECKED_USE_RELATIVE_PATH);
                MessageBox.Show("载入配置文件成功", "恭喜", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool _CheckConfig()
        {
            // 检查工具所在路径是否填写正确
            string programPath = tbProgramPath.Text.Trim();
            if (string.IsNullOrEmpty(programPath))
            {
                MessageBox.Show("必须指定XlsxToLua工具所在路径", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (programPath.Contains(AppValues.GUI_PROGRAM_NAME))
            {
                MessageBox.Show(string.Format("需要指定的是{0}所在路径而不是{1}所在路径", AppValues.PROGRAM_NAME, AppValues.GUI_PROGRAM_NAME), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!File.Exists(programPath))
            {
                MessageBox.Show("指定的XlsxToLua工具所在路径不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!".exe".Equals(Path.GetExtension(programPath), StringComparison.CurrentCultureIgnoreCase))
            {
                MessageBox.Show("指定的XlsxToLua工具错误，不是一个有效的exe程序", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            // 检查Excel文件所在目录是否填写正确
            string excelFolderPath = tbExcelFolderPath.Text.Trim();
            if (string.IsNullOrEmpty(excelFolderPath))
            {
                MessageBox.Show("必须指定Excel文件所在目录", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!Directory.Exists(excelFolderPath))
            {
                MessageBox.Show("指定的Excel文件所在目录不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            // 检查Excel文件夹及其下属子文件夹中是否存在同名文件
            // 记录Excel文件夹及其子文件夹中的文件名对应的所在路径（key：表名， value：文件所在路径）
            Dictionary<string, string> tableNameAndPath = new Dictionary<string, string>();
            // 记录重名文件所在目录
            Dictionary<string, List<string>> sameExcelNameInfo = new Dictionary<string, List<string>>();
            if (cbExportIncludeSubfolder.Checked == true)
            {
                foreach (string filePath in Directory.GetFiles(excelFolderPath, "*.xlsx", SearchOption.AllDirectories))
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    if (fileName.StartsWith(AppValues.EXCEL_TEMP_FILE_FILE_NAME_START_STRING))
                        continue;

                    if (tableNameAndPath.ContainsKey(fileName))
                    {
                        if (!sameExcelNameInfo.ContainsKey(fileName))
                        {
                            sameExcelNameInfo.Add(fileName, new List<string>());
                            sameExcelNameInfo[fileName].Add(tableNameAndPath[fileName]);
                        }

                        sameExcelNameInfo[fileName].Add(filePath);
                    }
                    else
                        tableNameAndPath.Add(fileName, filePath);
                }

                if (sameExcelNameInfo.Count > 0)
                {
                    StringBuilder sameExcelNameErrorStringBuilder = new StringBuilder();
                    sameExcelNameErrorStringBuilder.AppendLine("错误：Excel文件夹及其子文件夹中不允许出现同名文件，重名文件如下：");
                    foreach (var item in sameExcelNameInfo)
                    {
                        string fileName = item.Key;
                        List<string> filePath = item.Value;
                        sameExcelNameErrorStringBuilder.AppendFormat("以下路径中存在同名文件（{0}）：\n", fileName);
                        foreach (string oneFilePath in filePath)
                            sameExcelNameErrorStringBuilder.AppendLine(oneFilePath);
                    }

                    MessageBox.Show(sameExcelNameErrorStringBuilder.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            else
            {
                foreach (string filePath in Directory.GetFiles(excelFolderPath, "*.xlsx", SearchOption.TopDirectoryOnly))
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    if (fileName.StartsWith(AppValues.EXCEL_TEMP_FILE_FILE_NAME_START_STRING))
                        continue;

                    tableNameAndPath.Add(fileName, filePath);
                }
            }
            // 检查如果设置了-exportKeepDirectoryStructure参数，是否也设置了-exportIncludeSubfolder参数
            if (cbExportKeepDirectoryStructure.Checked == true && cbExportIncludeSubfolder.Checked == false)
            {
                MessageBox.Show(string.Format("只有通过设置{0}参数，将要导出的Excel文件夹下的各级子文件夹中的Excel文件也进行导出时，指定{1}参数设置将生成的文件按原Excel文件所在的目录结构进行存储才有意义，请检查是否遗漏设置{0}参数", AppValues.EXPORT_INCLUDE_SUBFOLDER_PARAM_STRING, AppValues.EXPORT_KEEP_DIRECTORY_STRUCTURE_PARAM_STRING), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            // 检查lua文件导出目录是否填写正确
            string exportLuaFolderPath = tbExportLuaFolderPath.Text.Trim();
            if (string.IsNullOrEmpty(exportLuaFolderPath))
            {
                MessageBox.Show(string.Format("必须指定lua文件导出目录（若不导出lua文件，请填写{0}）", AppValues.NOT_EXPORT_LUA_PARAM_STRING), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbExportLuaFolderPath.Text = AppValues.NOT_EXPORT_LUA_PARAM_STRING;
                return false;
            }
            if (!exportLuaFolderPath.Equals(AppValues.NOT_EXPORT_LUA_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase))
            {
                if (!Directory.Exists(exportLuaFolderPath))
                {
                    MessageBox.Show("指定的lua文件导出目录不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            // 检查项目Client所在目录是否填写正确
            string clientFolderPath = tbClientFolderPath.Text.Trim();
            if (string.IsNullOrEmpty(clientFolderPath))
            {
                MessageBox.Show(string.Format("未指定Client所在目录（若无需指定，请填写{0}）", AppValues.NO_CLIENT_PATH_PARAM_STRING), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbClientFolderPath.Text = AppValues.NO_CLIENT_PATH_PARAM_STRING;
                return false;
            }
            if (!clientFolderPath.Equals(AppValues.NO_CLIENT_PATH_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase))
            {
                if (!Directory.Exists(clientFolderPath))
                {
                    MessageBox.Show("指定的项目Client所在目录不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            // 检查lang文件所在路径是否填写正确
            string langFilePath = tbLangFilePath.Text.Trim();
            if (string.IsNullOrEmpty(langFilePath))
            {
                MessageBox.Show(string.Format("未指定lang文件所在路径（若无需指定，请填写{0}）", AppValues.NO_LANG_PARAM_STRING), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbLangFilePath.Text = AppValues.NO_LANG_PARAM_STRING;
                return false;
            }
            if (!langFilePath.Equals(AppValues.NO_LANG_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase))
            {
                if (!File.Exists(langFilePath))
                {
                    MessageBox.Show("指定的lang文件所在路径不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            // 若设置导出部分Excel文件，检查文件名声明是否正确
            List<string> exportTableNames = new List<string>();
            if (cbPart.Checked == true)
            {
                string partExcelNames = tbPartExcelNames.Text.Trim();
                if (string.IsNullOrEmpty(partExcelNames))
                {
                    MessageBox.Show("勾选了导出部分Excel文件的选项，就必须在文本框中填写要导出的Excel文件名", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                string[] fileNames = partExcelNames.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string fileName in fileNames)
                    exportTableNames.Add(fileName.Trim());

                // 检查指定导出的Excel文件是否存在
                foreach (string exportExcelFileName in exportTableNames)
                {
                    if (!tableNameAndPath.ContainsKey(exportExcelFileName))
                    {
                        MessageBox.Show(string.Format("指定要导出的Excel文件（{0}）不存在，请检查后重试并注意区分大小写", string.Concat(exportExcelFileName, ".xlsx")), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            // 若设置忽略导出部分Excel文件，检查文件名声明是否正确
            List<string> exceptTableNames = new List<string>();
            if (cbExcept.Checked == true)
            {
                string exceptExcelNames = tbExceptExcelNames.Text.Trim();
                if (string.IsNullOrEmpty(exceptExcelNames))
                {
                    MessageBox.Show("勾选了忽略导出部分Excel文件的选项，就必须在文本框中填写要忽略的Excel文件名", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                string[] fileNames = exceptExcelNames.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string fileName in fileNames)
                    exceptTableNames.Add(fileName.Trim());

                // 检查指定忽略导出的Excel文件是否存在
                foreach (string exceptTableName in exceptTableNames)
                {
                    if (!tableNameAndPath.ContainsKey(exceptTableName))
                    {
                        MessageBox.Show(string.Format("指定要忽略导出的Excel文件（{0}）不存在，请检查后重试并注意区分大小写", string.Concat(exceptTableName, ".xlsx")), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            // 同一张表格不能既设置为-part又设置为-except
            foreach (string exportTableName in exportTableNames)
            {
                if (exceptTableNames.Contains(exportTableName))
                {
                    MessageBox.Show(string.Format("对表格{0}既设置了-part参数，又设置了-except参数", exportTableNames), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;
        }

        private string _GetExecuteParamString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            Uri programUri = new Uri(AppValues.PROGRAM_PATH);

            string programPath = tbProgramPath.Text.Trim();
            string excelFolderPath = tbExcelFolderPath.Text.Trim();
            string exportLuaFolderPath = tbExportLuaFolderPath.Text.Trim();
            string clientFolderPath = tbClientFolderPath.Text.Trim();
            string langFilePath = tbLangFilePath.Text.Trim();
            if (cbIsUseRelativePath.Checked == true)
            {
                if (Path.IsPathRooted(programPath))
                    programPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(programPath)).ToString());
                if (Path.IsPathRooted(excelFolderPath))
                    excelFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(excelFolderPath)).ToString());
                if (!exportLuaFolderPath.Equals(AppValues.NOT_EXPORT_LUA_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(exportLuaFolderPath))
                    exportLuaFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(exportLuaFolderPath)).ToString());
                if (!clientFolderPath.Equals(AppValues.NO_CLIENT_PATH_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(clientFolderPath))
                    clientFolderPath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(clientFolderPath)).ToString());
                if (!langFilePath.Equals(AppValues.NO_LANG_PARAM_STRING, StringComparison.CurrentCultureIgnoreCase) && Path.IsPathRooted(langFilePath))
                    langFilePath = Uri.UnescapeDataString(programUri.MakeRelativeUri(new Uri(langFilePath)).ToString());
            }
            stringBuilder.AppendFormat("\"{0}\" ", programPath).AppendFormat("\"{0}\" ", excelFolderPath).AppendFormat("\"{0}\" ", exportLuaFolderPath).AppendFormat("\"{0}\" ", clientFolderPath).AppendFormat("\"{0}\" ", langFilePath);
            if (cbExportIncludeSubfolder.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.EXPORT_INCLUDE_SUBFOLDER_PARAM_STRING);
            if (cbExportKeepDirectoryStructure.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.EXPORT_KEEP_DIRECTORY_STRUCTURE_PARAM_STRING);
            if (cbColumnInfo.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.NEED_COLUMN_INFO_PARAM_STRING);
            if (cbUnchecked.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.UNCHECKED_PARAM_STRING);
            if (cbPrintEmptyStringWhenLangNotMatching.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.LANG_NOT_MATCHING_PRINT_PARAM_STRING);
            if (cbExportMySQL.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.EXPORT_MYSQL_PARAM_STRING);
            if (cbAllowedNullNumber.Checked == true)
                stringBuilder.AppendFormat("\"{0}\" ", AppValues.ALLOWED_NULL_NUMBER_PARAM_STRING);
            if (cbPart.Checked == true)
            {
                string partExcelNames = tbPartExcelNames.Text.Trim();
                stringBuilder.AppendFormat("\"{0}({1})\" ", AppValues.PART_EXPORT_PARAM_STRING, partExcelNames);
            }
            if (cbExcept.Checked == true)
            {
                string exceptExcelNames = tbExceptExcelNames.Text.Trim();
                stringBuilder.AppendFormat("\"{0}({1})\" ", AppValues.EXCEPT_EXPORT_PARAM_STRING, exceptExcelNames);
            }

            return stringBuilder.ToString();
        }

        private void _WarnWhenChooseDangerousParam(CheckBox cb)
        {
            if (cb.Checked == true)
                cb.ForeColor = Color.Red;
            else
                cb.ForeColor = Color.Black;
        }

        /// <summary>
        /// 弹出文件选择对话框，选择部分要进行操作的Excel文件
        /// </summary>
        private string _GetChoosePartExcelFile(string originalText, out string errorString)
        {
            string excelFolderPath = tbExcelFolderPath.Text.Trim();
            if (string.IsNullOrEmpty(excelFolderPath))
            {
                errorString = "请先指定Excel文件所在目录";
                return null;
            }
            if (!Directory.Exists(excelFolderPath))
            {
                errorString = "指定Excel文件所在目录不存在，请重新设置";
                return null;
            }

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "请选择Excel表格";
            dialog.InitialDirectory = excelFolderPath;
            dialog.Multiselect = true;
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string[] filePaths = dialog.FileNames;
                // 检查选择的Excel文件是否在设置的Excel所在目录
                string checkFilePath = Path.GetDirectoryName(filePaths[0]);
                if (cbExportIncludeSubfolder.Checked == true)
                {
                    if (!Path.GetFullPath(checkFilePath).StartsWith(Path.GetFullPath(excelFolderPath), StringComparison.CurrentCultureIgnoreCase))
                    {
                        errorString = string.Format("必须在指定的Excel文件所在目录或子目录中选择导出文件\n设置的Excel文件所在根目录为：{0}\n而你选择的Excel所在目录为：{1}", Path.GetFullPath(excelFolderPath), Path.GetFullPath(checkFilePath));
                        return null;
                    }
                }
                else
                {
                    if (!Path.GetFullPath(checkFilePath).Equals(Path.GetFullPath(excelFolderPath), StringComparison.CurrentCultureIgnoreCase))
                    {
                        errorString = string.Format("必须在指定的Excel文件所在目录中选择导出文件\n设置的Excel文件所在目录为：{0}\n而你选择的Excel所在目录为：{1}", Path.GetFullPath(excelFolderPath), Path.GetFullPath(checkFilePath));
                        return null;
                    }
                }

                List<string> fileNames = new List<string>();
                foreach (string filePath in filePaths)
                    fileNames.Add(Path.GetFileNameWithoutExtension(filePath));

                //originalText = originalText.Trim();
                //if (!string.IsNullOrEmpty(originalText))
                //{
                //    string[] originalInputExcelFile = originalText.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                //    for (int i = 0; i < originalInputExcelFile.Length; ++i)
                //    {
                //        string oneOriginalInputExcelFile = originalInputExcelFile[i].Trim();
                //        if (!string.IsNullOrEmpty(oneOriginalInputExcelFile) && !fileNames.Contains(oneOriginalInputExcelFile))
                //            fileNames.Add(oneOriginalInputExcelFile);
                //    }
                //}

                errorString = null;
                return Utils.CombineString(fileNames, "|");
            }
            else
            {
                errorString = null;
                return null;
            }
        }

        /// <summary>
        /// 想拖拽接收文件或文件夹路径的文本框，需要注册DragEnter事件回调
        /// </summary>
        private void _TextBoxDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Link;
        }

        /// <summary>
        /// 要求拖拽一个文件到文本框的DragDrop事件回调
        /// </summary>
        private void _TextBoxOneFileDragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) == true)
            {
                Array dragDropFileArray = e.Data.GetData(DataFormats.FileDrop) as Array;
                if (dragDropFileArray.Length != 1)
                {
                    MessageBox.Show("只允许拖入一个指定的文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string path = dragDropFileArray.GetValue(0).ToString();
                if (Directory.Exists(path) == true)
                {
                    MessageBox.Show("请拖入一个指定的文件而不是文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    if (File.Exists(path))
                    {
                        TextBox textBox = sender as TextBox;
                        textBox.Text = path;
                    }
                    else
                    {
                        MessageBox.Show("请拖入一个指定的文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("请拖入一个指定的文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// 要求拖拽一个文件夹到文本框的DragDrop事件回调
        /// </summary>
        private void _TextBoxOneDirDragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) == true)
            {
                Array dragDropDirArray = e.Data.GetData(DataFormats.FileDrop) as Array;
                if (dragDropDirArray.Length != 1)
                {
                    MessageBox.Show("只允许拖入一个指定的文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string path = dragDropDirArray.GetValue(0).ToString();
                if (File.Exists(path) == true)
                {
                    MessageBox.Show("请拖入一个指定的文件夹而不是文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    if (Directory.Exists(path))
                    {
                        TextBox textBox = sender as TextBox;
                        textBox.Text = path;
                    }
                    else
                    {
                        MessageBox.Show("请拖入一个指定的文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("请拖入一个指定的文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// 在Excel所在目录中选择多个表格文件拖拽到文本框的DragDrop事件回调
        /// </summary>
        private void _TextBoxMultipleExcelFileDragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) == true)
            {
                string excelFolderPath = tbExcelFolderPath.Text.Trim();
                if (string.IsNullOrEmpty(excelFolderPath))
                {
                    MessageBox.Show("请先指定Excel文件所在目录", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!Directory.Exists(excelFolderPath))
                {
                    MessageBox.Show("指定Excel文件所在目录不存在，请重新设置", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Array dragDropFileArray = e.Data.GetData(DataFormats.FileDrop) as Array;
                List<string> fileNames = new List<string>();
                foreach (string path in dragDropFileArray)
                {
                    if (Directory.Exists(path) == true)
                    {
                        MessageBox.Show("请在Excel所在文件夹中选择并拖入若干个表格文件，而不是文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (File.Exists(path) == false)
                    {
                        MessageBox.Show("请在Excel所在文件夹中选择并拖入若干个表格文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    // 检查选择的Excel文件是否在设置的Excel所在目录
                    string checkFilePath = Path.GetDirectoryName(path);
                    if (cbExportIncludeSubfolder.Checked == true)
                    {
                        if (!Path.GetFullPath(checkFilePath).StartsWith(Path.GetFullPath(excelFolderPath), StringComparison.CurrentCultureIgnoreCase))
                        {
                            string errorString = string.Format("必须在指定的Excel文件所在目录或子目录中选择导出文件\n设置的Excel文件所在根目录为：{0}\n而你选择的Excel所在目录为：{1}", Path.GetFullPath(excelFolderPath), Path.GetFullPath(checkFilePath));
                            MessageBox.Show(errorString, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else
                    {
                        if (!Path.GetFullPath(checkFilePath).Equals(Path.GetFullPath(excelFolderPath), StringComparison.CurrentCultureIgnoreCase))
                        {
                            string errorString = string.Format("必须在指定的Excel文件所在目录中选择导出文件\n设置的Excel文件所在目录为：{0}\n而你选择的Excel所在目录为：{1}", Path.GetFullPath(excelFolderPath), Path.GetFullPath(checkFilePath));
                            MessageBox.Show(errorString, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    // 检查是否为Excel文件
                    if (!".xlsx".Equals(Path.GetExtension(path), StringComparison.CurrentCultureIgnoreCase))
                    {
                        MessageBox.Show("选择的必须都是Excel文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    fileNames.Add(Path.GetFileNameWithoutExtension(path));
                }

                TextBox textBox = sender as TextBox;
                textBox.Text = Utils.CombineString(fileNames, "|");
            }
            else
            {
                MessageBox.Show("请在Excel所在文件夹中选择并拖入若干个表格文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void cbExportJson_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
