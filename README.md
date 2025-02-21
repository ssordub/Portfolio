# System Configuration Utilities

A collection of professional-grade system configuration and management utilities built with Python and PowerShell. This toolkit provides a comprehensive GUI interface and batch scripts for Windows system administration tasks.

## 🛠️ Components

### System Configuration Utility (staging_setup_utility.py)

A feature-rich GUI application built with Python/Tkinter that provides a unified interface for various system administration tasks:

- **Hardware Management**
  - Scan and display all connected hardware devices
  - Export hardware information to JSON format
  - Sort and filter device listings
  - Automatic device categorization

- **Network Configuration**
  - DHCP configuration
  - Static IP setup with validation
  - DNS server management
  - Network interface configuration

- **System Setup**
  - Computer name management
  - Time zone configuration
  - Windows activation status and management
  - System environment variable configuration

- **File Management**
  - Dual-pane file transfer interface
  - Drive selection and navigation
  - File operations (copy, move, rename)
  - Context menu support

### Batch Utilities

#### ChangeEnvironmentalVariables.bat
- PowerShell-based environment variable manager
- Interactive command-line interface
- System-wide variable modification
- Secure value handling

#### GetHardwareDevices.bat
- Hardware device enumeration script
- Detailed device information reporting
- Manufacturer and device ID tracking
- User-friendly output formatting

## 🚀 Features

- **Modern GUI Interface**: Clean, professional design with intuitive navigation
- **Error Handling**: Comprehensive error catching and user feedback
- **PowerShell Integration**: Seamless execution of system commands
- **Cross-Component Communication**: Unified workflow between GUI and batch utilities
- **Security Conscious**: Proper permission handling and user confirmation for system changes

## 💻 Technologies

- Python 3.x
- Tkinter
- PowerShell
- Windows Management Instrumentation (WMI)
- Win32API

## 🛡️ Security Features

- User confirmation for system modifications
- Permission validation
- Secure command execution
- Input validation
- Error logging and handling

## 🔧 Installation

1. Clone the repository:
```bash
git clone https://github.com/ssordub/Portfolio.git
```

2. Ensure Python 3.X is installed on your system

3. Install required Python packages:
```bash
pip install tkinter
```

4. Run the main utility:
```bash
python staging_setup_utility.py
```

## 📚 Usage

### GUI Utility
0. (Optional: Use PyInstaller to create an executable version)
1. Launch `staging_setup_utility.py`
2. Navigate through the tabs for different functionality:
   - Hardware Devices
   - Network Setup
   - System Setup
   - Transfer Files

### Powershell/Batch Scripts

- Run `ChangeEnvironmentalVariables.bat` as administrator to modify system environment variables
- Execute `GetHardwareDevices.bat` to get a detailed hardware report
- Run `SimpleStagingSetup.bat` as administrator to configure a new device after imaging

## 📫 Contact

Your Name - bdon.ross@gmail.com

Portfolio Link: [https://github.com/ssordub/Portfolio](https://github.com/yourusername/system-utilities)
