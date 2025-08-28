# 邮件导出工具打包指南

本文档介绍如何将邮件导出工具打包成Windows可执行文件和安装程序。

## 📋 前置要求

### 必需软件
1. **Python 3.7+** - 确保已安装Python
2. **PyInstaller** - 用于打包Python程序
3. **Inno Setup** - 用于创建Windows安装程序

### 安装依赖

```bash
# 安装PyInstaller
pip install pyinstaller

# 安装项目依赖
pip install -r requirements.txt
```

### 下载Inno Setup

从官网下载并安装Inno Setup: https://jrsoftware.org/isinfo.php

## 🚀 快速打包

### 方法一：使用批处理脚本（推荐）

```bash
# 运行自动打包脚本
build.bat
```

这个脚本会自动：
- 检查依赖环境
- 清理之前的构建文件
- 使用PyInstaller打包程序
- 使用Inno Setup创建安装程序

### 方法二：使用Python配置脚本

```bash
# 使用Python打包脚本
python build_config.py
```

这个脚本提供更详细的配置选项和错误处理。

### 方法三：手动打包

#### 步骤1：打包可执行文件

```bash
# 打包GUI程序
pyinstaller --onefile --windowed --name="mail_exporter_gui" mail_exporter_gui.py

# 打包命令行程序
pyinstaller --onefile --name="mail_exporter" mail_exporter.py
```

#### 步骤2：创建安装程序

```bash
# 使用Inno Setup编译器
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
```

## 📁 文件结构

打包相关文件说明：

```
mail_exporter/
├── build.bat              # 自动打包批处理脚本
├── build_config.py        # Python打包配置脚本
├── setup.iss              # Inno Setup配置文件
├── icon.svg               # 应用程序图标（SVG格式）
├── BUILD.md               # 打包说明文档（本文件）
├── dist/                  # 打包输出目录
│   ├── mail_exporter_gui.exe
│   ├── mail_exporter.exe
│   └── 邮件导出工具_v1.0.0_安装程序.exe
└── build/                 # 临时构建文件
```

## ⚙️ 配置说明

### PyInstaller配置

`build_config.py`中的主要配置选项：

- `--onefile`: 打包成单个可执行文件
- `--windowed`: GUI程序不显示控制台窗口
- `--console`: CLI程序显示控制台窗口
- `--exclude-module`: 排除不需要的模块以减小文件大小
- `--optimize=2`: 优化字节码
- `--strip`: 移除调试信息

### Inno Setup配置

`setup.iss`中的主要配置：

- **应用信息**: 名称、版本、发布者等
- **安装选项**: 安装目录、权限要求等
- **文件包含**: 指定要包含的文件和目录
- **快捷方式**: 桌面和开始菜单快捷方式
- **卸载处理**: 自动检测和卸载旧版本

## 🔧 自定义配置

### 修改应用信息

编辑`setup.iss`文件顶部的定义：

```pascal
#define MyAppName "邮件导出工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "你的名字"
#define MyAppURL "https://github.com/your-username/mail-exporter"
```

### 添加图标

1. 将SVG图标转换为ICO格式：
   ```bash
   # 使用在线工具或ImageMagick
   magick icon.svg icon.ico
   ```

2. 更新`setup.iss`中的图标路径：
   ```pascal
   SetupIconFile=icon.ico
   ```

### 优化文件大小

在`build_config.py`中添加更多排除模块：

```python
'--exclude-module=模块名',
```

## 🐛 常见问题

### 问题1：PyInstaller打包失败

**解决方案**：
- 确保所有依赖都已安装
- 检查Python路径是否正确
- 尝试在虚拟环境中打包

### 问题2：程序运行时缺少模块

**解决方案**：
- 在PyInstaller命令中添加`--hidden-import=模块名`
- 检查`build_config.py`中的隐藏导入配置

### 问题3：Inno Setup编译失败

**解决方案**：
- 确保Inno Setup已正确安装
- 检查`setup.iss`中的文件路径
- 确保`dist`目录中存在所需的可执行文件

### 问题4：安装程序无法运行

**解决方案**：
- 检查目标系统的架构（x86/x64）
- 确保目标系统有必要的运行时库
- 以管理员权限运行安装程序

## 📊 打包优化建议

### 减小文件大小

1. **排除不必要的模块**：
   ```python
   '--exclude-module=matplotlib',
   '--exclude-module=numpy',
   '--exclude-module=pandas',
   ```

2. **使用UPX压缩**（可选）：
   ```bash
   # 下载UPX并添加到PATH
   pyinstaller --upx-dir=path/to/upx ...
   ```

3. **优化导入**：
   - 只导入需要的模块
   - 避免导入大型库的整个模块

### 提高兼容性

1. **指定Python版本**：
   - 使用较旧的Python版本以提高兼容性
   - 在虚拟环境中使用特定版本

2. **测试多个系统**：
   - Windows 7/8/10/11
   - 32位和64位系统

## 📝 版本发布流程

1. **更新版本号**：
   - 修改`setup.iss`中的版本号
   - 更新`README.md`中的版本信息

2. **测试打包**：
   ```bash
   python build_config.py
   ```

3. **测试安装程序**：
   - 在干净的系统上测试安装
   - 验证所有功能正常工作

4. **发布**：
   - 上传到GitHub Releases
   - 提供安装说明

## 📞 技术支持

如果在打包过程中遇到问题，请：

1. 检查本文档的常见问题部分
2. 查看PyInstaller和Inno Setup的官方文档
3. 在项目仓库中提交Issue

---

**注意**: 首次打包可能需要较长时间，后续打包会更快。建议在发布前在多个系统上测试安装程序。