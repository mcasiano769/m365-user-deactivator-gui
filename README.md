# m365-user-deactivator-gui

A user-friendly Python GUI tool for automated Microsoft 365 user offboarding and deactivation.

![Python](https://img.shields.io/badge/python-v3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-windows-lightgrey.svg)

## 🚀 Features

- **Convert mailbox to shared** - Prepares user mailbox for shared access
- **Remove and log licenses** - Removes all M365 licenses and saves a record
- **Block sign-in** - Prevents user from accessing M365 services
- **Reset password** - Sets password to `360Rules!`
- **Revoke Intune sessions** - Finds and revokes all managed device sessions
- **Reset MFA devices** - Removes registered multi-factor authentication methods
- **Automated reporting** - Generates detailed completion report on Desktop
- **User-friendly GUI** - Simple interface requiring only first and last name

## 📋 Requirements

- **Python 3.8+**
- **Windows OS** (for Desktop report saving)
- **Microsoft 365 Admin Access**
- **Azure App Registration** with appropriate permissions

## 🛠️ Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/m365-user-deactivator-gui.git
cd m365-user-deactivator-gui
