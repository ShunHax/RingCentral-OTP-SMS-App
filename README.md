# OTP Texting App

** This app has been largely built by use of ChatGPT 4o. I'm 100% sure it can be made more efficient/smaller. I'm not interested in continuing with this app, so here ya go. **

The **OTP Texting App** is a streamlined solution for generating and sending One-Time Passwords (OTPs) securely. Built using .NET 8, this app combines simplicity with powerful functionality, providing an executable-ready application for deployment.

## Features
- **Secure OTP Generation:** Dynamically create secure OTPs for authentication.
- **Customizable Interface:** Intuitive UI with warning and error icons for enhanced user interaction.
- **RingCentral Integration:** Leverage RingCentral for SMS delivery.
- **Cross-Language Support:** Resources included for multiple languages.

## Requirements
- **Operating System:** Windows (x64 and x86).
- **Framework:** .NET 8 Runtime.
- **Dependencies:** 
  - RingCentral.Net
  - Newtonsoft.Json
  - Microsoft.Web.WebView2

## Getting Started

### Installation
1. **Download the Executable:**
   - Navigate to the root directory and locate `OTP Texting App_Production.zip`.
   - Download the file and extract it.
   - Copy the executable to your desired location.

2. **Install .NET 8 Runtime:**
   - Ensure the required .NET 8 runtime is installed. [Download .NET Runtime](https://dotnet.microsoft.com/download).

### Running the Application
1. Double-click `OTP Texting App.exe` to launch the app.
2. Follow the on-screen prompts to configure and send OTPs.
   *Files are created in your %temp% directory, as well as the log.*

### Publishing the Application
For creating a new build:
1. Open the project in Visual Studio 2022.
2. Select `Release` configuration.
3. Build the project (`Ctrl + Shift + B`).
4. The built executable will be in the `bin\Release\net8.0-windows\win-x64\` directory.

## File Structure
```
OTP Texting App/
├── bin/ (Ignored in Git)
├── dist/
│   ├── OTP Texting App.exe
│   ├── Resources/ (Icons and other assets)
├── obj/ (Ignored in Git)
├── My Project/
│   ├── Application.Designer.vb
│   ├── PublishProfiles/ (Ignored in Git)
├── README.md
├── .gitignore
└── OTP Texting App.sln
```

## Contributing
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/otp-texting-app.git
   ```
2. Create a new branch:
   ```bash
   git checkout -b feature/your-feature
   ```
3. Commit changes:
   ```bash
   git commit -m "Add your feature"
   ```
4. Push to GitHub:
   ```bash
   git push origin feature/your-feature
   ```
5. Create a Pull Request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgments
- **Microsoft** for .NET 8.
- **RingCentral** for SMS APIs.
- **Newtonsoft** for JSON parsing.

---
### Notes
- Always ensure the `dist/` directory contains the latest executable when publishing.
- For advanced deployment, consider signing the executable to avoid security warnings.
