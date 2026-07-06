# 🧰 工具箱 Toolbox

![Toolbox Icon](toolbox_icon.png)

一个集成多种实用工具的桌面应用程序，提高工作效率，简化重复操作。

## ✨ 功能特点

### 📹 功能1：视频批量重命名
支持抖音和视频号两种整理方式，快速批量重命名视频文件。

**功能亮点：**
- 📊 支持导入xlsx、xls、csv格式表格
- 🔍 自动筛选消耗值 > 1500 的数据
- 📝 智能生成规范化视频命名
- ↩️ 支持一键撤销重命名操作
- 📋 实时操作日志显示

**使用方法：**
1. 选择整理方式（抖音/视频号）
2. 导入原始数据表格
3. 点击"整理"生成命名规则
4. 选择视频文件夹执行重命名
5. 如需撤销，点击"撤销重命名"

---

### 🔍 功能2：筛选未绑账户
快速筛选出未绑定账户的数据，支持一键复制到剪贴板。

**功能亮点：**
- ⚡ 快速筛选未绑账户
- 📋 支持一键复制到剪贴板
- 📊 实时显示筛选结果

---

### 📝 功能3：模板自动填写
合同模板自动填充，提高合同制作效率。

**功能亮点：**
- 📄 支持Word文档模板（.docx）
- 🔍 自动扫描占位符（如{{date}}、{{company}}等）
- 📅 智能日期计算（合规签约日、到期日期）
- 💾 模板配置永久保存
- 🔄 支持批量更新模板路径
- 🎨 保留原格式（字体、颜色、下划线）

**占位符说明：**
- `{{date}}` - 当前日期（YYYY年MM月DD日）
- `{{company}}` - 主体名称（手动填写）
- `{{Date_of_Signing}}` - 合规签约日（1日/15日/月末）
- `{{out_of_date}}` - 到期日期（签订日期+1年-1天）

---

### 🔢 功能4：数值转换
快速转换数值格式，处理常规格式单元格。

## 📥 安装使用

### 方法1：直接使用EXE文件（推荐）

1. 下载 `Toolbox.exe` 文件
2. 双击运行即可使用（无需安装Python环境）

### 方法2：Python环境运行

#### 系统要求：
- Python 3.8 或更高版本
- Windows 操作系统

#### 安装步骤：

```bash
# 1. 克隆仓库
git clone https://github.com/xiaobaiNINiZ/Toolbox.git

# 2. 进入项目目录
cd Toolbox

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
python video_rename_tool.py
```

#### 依赖库：
- `tkinter` - GUI界面
- `openpyxl` - Excel文件处理
- `python-docx` - Word文档处理
- `Pillow` - 图标处理

## 🛠️ 技术栈

- **开发语言**: Python 3.14
- **GUI框架**: Tkinter
- **打包工具**: PyInstaller
- **文件处理**: openpyxl, python-docx
- **图像处理**: Pillow

## 📁 项目结构

```
Toolbox/
├── video_rename_tool.py    # 主程序代码
├── toolbox_icon.ico        # 工具箱图标
├── toolbox_icon.png        # 图标PNG版本
├── Toolbox.spec            # PyInstaller打包配置
├── requirements.txt        # Python依赖列表
├── build_exe.bat           # 打包脚本
├── templates.json          # 模板配置文件（运行时生成）
├── README.md               # 项目说明文档
└── .gitignore              # Git忽略配置
```

## 🎯 使用截图

程序界面包含4个功能标签页：
- 视频批量重命名
- 筛选未绑账户
- 模板自动填写
- 数值转换

## 📝 更新日志

### v1.0.0 (2026-07-03)
- ✨ 初始版本发布
- 📹 实现视频批量重命名功能（支持抖音和视频号）
- 🔍 实现筛选未绑账户功能
- 📝 实现模板自动填写功能
- 🔢 实现数值转换功能
- 🎨 添加工具箱图标
- 📦 打包为独立EXE文件

## 🤝 贡献指南

欢迎提出问题和改进建议！

1. Fork 本仓库
2. 创建新分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

## 📄 许可证

本项目采用 MIT 许可证。

## 👨‍💻 作者

- **GitHub**: [@xiaobaiNINiZ](https://github.com/xiaobaiNINiZ)
- **邮箱**: 1480363335@qq.com

## 🙏 致谢

感谢所有开源项目的贡献者，让这个工具得以实现！

---

⭐ 如果这个项目对您有帮助，请给个星标支持！