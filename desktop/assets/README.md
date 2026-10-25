# EasyExcel 应用图标

0.6.1 起使用专属绿色 Logo：表格单元格构成字母 E，搭配薄荷绿强调色，适配软件现有的绿色界面。圆角方形外为真实透明背景，便于在不同桌面主题中显示。

- `logo-source.png`：内置 imagegen 生成并清理透明边缘的高清 PNG 原稿。
- `app.png`：512 × 512，用于界面品牌标识、网页图标与非 Windows 窗口。
- `app.ico`：16、20、24、32、40、48、64、128、256 像素的 32 位图标，用于 Windows 可执行文件、窗口和快捷方式。
- `logo-prompt.txt`：生成与边缘清理时使用的完整提示词。使用内置 imagegen，没有调用 CLI 或单独配置 API Key。

原稿确定后，使用 Windows 自带图形库缩放和转换格式，保留透明通道；默认完整保留方形图片，非方形图片居中裁切：

```powershell
.\desktop\scripts\make-app-icon.ps1 -Source '.\desktop\assets\logo-source.png'
```

如果需要转换其他构图，可显式指定 `-CropX`、`-CropY`、`-CropSize`。软件各处统一引用 `app.png` / `app.ico`，修改后重新构建免安装包即可。
