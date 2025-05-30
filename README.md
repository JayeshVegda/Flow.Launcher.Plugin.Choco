# Flow.Launcher.Plugin.Choco

A powerful [Flow Launcher](https://github.com/Flow-Launcher/Flow.Launcher) plugin for managing Chocolatey packages directly from your desktop. This plugin provides a seamless interface to search, install, uninstall, and manage your Chocolatey packages without leaving your keyboard.

![Plugin Preview](Images/preview.png)

## Features

- 🔍 **Quick Search**: Instantly search through available Chocolatey packages
- 📦 **Package Management**: 
  - Install new packages
  - Uninstall existing packages
  - Upgrade installed packages
- 📋 **Installed Packages List**: View all your installed packages at a glance
- ⚡ **Caching System**: Efficient package list caching for faster response times
- 🔄 **Real-time Updates**: Cache automatically updates after installations/uninstallations
- 📊 **Detailed Package Information**: View package descriptions, download counts, and tags

## Installation

1. Open Flow Launcher
2. Press `Ctrl+Shift+P` to open the plugin manager
3. Search for "Chocolatey"
4. Click "Install" on the Flow.Launcher.Plugin.Choco plugin

## Usage

### Basic Commands

- **Empty Query**: Shows list of installed packages
- **Search**: Type any text to search for packages
- **Install**: Press Enter on a search result to install
- **Uninstall**: Press Enter on an installed package to uninstall

### Keyboard Shortcuts

- `Enter`: Install/Uninstall selected package
- `Esc`: Close Flow Launcher
- `Up/Down`: Navigate through results

## Requirements

- Windows OS
- [Flow Launcher](https://github.com/Flow-Launcher/Flow.Launcher) installed
- [Chocolatey](https://chocolatey.org/install) package manager installed
- PowerShell 7 or later

## Features in Detail

### Package Search
- Real-time search through Chocolatey repository
- Detailed package information including:
  - Version
  - Description
  - Download count
  - Tags

### Package Management
- One-click installation
- Easy uninstallation
- Automatic cache updates
- Administrative privileges handling

### Caching System
- Efficient JSON-based caching
- Automatic cache invalidation
- Real-time cache updates
- Improved performance for repeated queries

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Flow Launcher](https://github.com/Flow-Launcher/Flow.Launcher) for the amazing launcher framework
- [Chocolatey](https://chocolatey.org/) for the package management system

## Support

If you encounter any issues or have suggestions, please:
1. Check the [Issues](https://github.com/yourusername/Flow.Launcher.Plugin.Choco/issues) page
2. Create a new issue if your problem isn't already listed

## Author

[Your Name] - [Your GitHub Profile]

---

⭐ Star this repository if you find it useful! 