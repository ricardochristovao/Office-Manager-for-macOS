# Office Manager for macOS

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-macOS-lightgrey.svg)
![Bash](https://img.shields.io/badge/bash-4.0%2B-brightgreen.svg)

A command-line utility for managing Microsoft Office 2024 installations on macOS systems.

## 🇧🇷 Descrição

Este script bash automatiza a gestão do Microsoft Office 2024 para macOS, permitindo:

- Detectar instalações existentes do Microsoft Office
- Remover versões incompatíveis (anteriores ao Office 2024)
- Instalar o Microsoft Office 2024
- Ativar o Office 2024 usando o serializador VL

## 🇺🇸 Description

This bash script automates Microsoft Office 2024 management for macOS, allowing users to:

- Detect existing Microsoft Office installations
- Remove incompatible versions (pre-Office 2024)
- Install Microsoft Office 2024
- Activate Office 2024 using the VL serializer

## 📋 Requisitos / Requirements

- macOS 11.0 (Big Sur) ou superior
- Permissões de administrador
- Conexão com a internet

## ⚙️ Instalação / Installation

```bash
# Clone o repositório
git clone https://github.com/ricardochristovao/office-manager.git

# Entre no diretório
cd office-manager

# Torne o script executável
chmod +x office_manager.sh

# Execute o script
./office_manager.sh
