# HashTool

[中文说明](README_CN.md)

HashTool is a lightweight **Windows desktop application** built with **VB.NET and WPF**, designed specifically for **file hash verification**. It currently supports **MD5** and **SHA256**, making it useful for checking file integrity and validating downloads.
![Uploading image.png…]()

## ✨ Features

* File hash verification only (no text hashing)
* Supports **MD5** and **SHA256**
* Simple and clean WPF user interface
* Native Windows experience
* Clear project structure, easy to read and extend

## 🖼️ Screenshot

> You can add screenshots of the application here

## 🛠️ Tech Stack

* **Language**: VB.NET
* **Framework**: WPF (.NET)
* **UI**: XAML
* **Platform**: Windows

## 📂 Project Structure

```text
├── MainWindow.xaml              # Main window UI
├── MainWindow.xaml.vb           # Main window logic
├── HashToolControl.xaml         # Hash tool control UI
├── HashToolControl.xaml.vb      # Hash tool control logic
├── DwmHelper.vb                 # Windows DWM helper
├── NativeMethods.vb             # Win32 API wrappers
└── README.md                    # Project documentation
```

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/luoji12/HashTool.git
```

### 2. Open in Visual Studio

* Visual Studio 2019 or later is recommended
* Make sure the required .NET workload is installed

### 3. Build and Run

* Build the solution
* Select a file in the UI to calculate its hash
* Compare the result with the expected MD5 or SHA256 value

## 📌 Requirements

* Windows 10 / 11
* .NET Framework or .NET (as defined in the project)
* Visual Studio (for development)

## 📖 Possible Improvements

You may extend this project by:

* Adding more hash algorithms (SHA1, SHA512, etc.)
* Drag-and-drop file support
* Batch file hashing
* Dark mode / theme customization
* Multilingual UI

## 🤝 Contributing

Contributions are welcome!

* Feel free to open Issues for bugs or feature requests
* Pull Requests are appreciated

## 📄 License

This project is licensed under the **MIT License**.

---

If you find this project useful, consider giving it a ⭐ on GitHub!
