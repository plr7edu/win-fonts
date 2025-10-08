# Windows Fonts Collection

A curated collection of fonts and utilities for Windows font management.

## 📦 Included Fonts

| Font Name | File | Format |
|-----------|------|--------|
| Cascadia Cove Nerd Font | `CaskaydiaCoveNerdFont.otf` | OTF |
| Cascadia Cove Nerd Font Mono | `CaskaydiaCoveNerdFontMono.otf` | OTF |
| Hack Nerd Font Mono | `HackNerdFontMono.ttc` | TTC |
| IBM Plex Sans | `IBMPlexSans.ttc` | TTC |
| JetBrains Mono Nerd Font | `JetBrainsMonoNerdFont.ttf` | TTF |
| JetBrains Mono Nerd Font Mono | `JetBrainsMonoNerdFontMono.ttf` | TTF |
| Noto Sans Mono Nerd Font | `NotoSansMonoNerdFont.ttc` | TTC |
| Roboto Slab | `RobotoSlab.ttf` | TTF |
| Space Grotesk | `SpaceGrotesk.ttc` | TTC |

## 🛠️ Tools

| Tool | Description |
|------|-------------|
| `1_Font_Installer.exe` | Install fonts to Windows |
| `2_Font_Uninstaller.exe` | Uninstall fonts from Windows |
| `!font_name_editor.exe` | Edit font metadata and names |
| `3-rename-font-pattern.vbs` | Batch rename fonts using patterns |

## 🐍 Python Font Tools

Located in `!python-font-tools/` directory:

| Tool | Purpose |
|------|---------|
| `font_disassembler` | Disassemble font files for analysis |
| `font_name_editor` | Edit font names programmatically |
| `one_family_fonts_assembler` | Combine fonts into a single family |

### Requirements

Install Python dependencies:

```bash
pip install -r !python-font-tools/!requirements/requirements.txt
```

## 📥 Installation

1. Run `1_Font_Installer.exe` to install fonts
2. Select the font files you want to install
3. Fonts will be installed system-wide

## 🗑️ Uninstallation

1. Run `2_Font_Uninstaller.exe` to remove fonts
2. Select the fonts you want to uninstall
3. Fonts will be removed from the system

## 📚 Source & Credits

The font installer/uninstaller tools (`0-font-install-uninstall-source.7z`) are based on [**FontRegister**](https://github.com/Nucs/FontRegister) by [@Nucs](https://github.com/Nucs).

**FontRegister** is a C# library and utility for programmatically installing and uninstalling fonts on Windows.

