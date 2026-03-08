# 物理实验不确定度计算工具

一个面向物理实验课的 Windows 桌面工具，用于完成 A 类评定、B 类评定、合成标准不确定度和扩展不确定度计算，并支持项目保存、Excel 导入、数据/结果 Excel 导出、TXT 导出和结果图片导出。

临时 Vibe Coding 出来的结果，顺便开源供大家改进和使用啦~ 

## 当前版本范围

- 单一物理量的不确定度计算
- A 类评定：重复测量值统计
- B 类评定：分度值、允差/准确度、厂家给定标准差、自定义分布因子
- 合成标准不确定度与扩展不确定度
- 项目文件保存与自动恢复
- Excel 导入与数据/结果 Excel、TXT、图片导出
- 结果显示支持自动修约或固定小数位
- 跟随系统深浅色模式切换界面主题
- 支持从 GitHub Release 检测更新、查看更新说明并跳转下载
- 可在程序设置中关闭启动时自动检查更新

## 安装依赖

建议使用虚拟环境进行开发。

```bash
pip install -r requirements.txt
```

## 启动

```bash
python main.py
```

## 项目文件

- 项目文件扩展名：`.uncx`
- 自动保存文件会写入系统应用数据目录

### 打包为 exe

先安装 PyInstaller：

```bash
pip install pyinstaller
```

然后在项目根目录执行：

```bash
pyinstaller --noconfirm --windowed --name 物理实验不确定度计算工具 --icon assets/app.ico main.py
```

打包完成后，可执行文件会出现在 `dist/物理实验不确定度计算工具/` 下。