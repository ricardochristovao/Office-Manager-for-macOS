#!/bin/bash

# Colors for better readability
GREEN='\033[0;32m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Function to check if any Office application is installed
check_office_installed() {
    echo -e "${BLUE}Verificando instalação do Office...${NC}"
    
    if [ -d "/Applications/Microsoft Word.app" ] || 
       [ -d "/Applications/Microsoft Excel.app" ] || 
       [ -d "/Applications/Microsoft PowerPoint.app" ] || 
       [ -d "/Applications/Microsoft Outlook.app" ]; then
        return 0 # Office is installed
    else
        return 1 # Office is not installed
    fi
}

# Function to identify Office version
identify_office_version() {
    if [ -d "/Applications/Microsoft Word.app" ]; then
        VERSION=$(defaults read "/Applications/Microsoft Word.app/Contents/Info" CFBundleShortVersionString 2>/dev/null)
        echo "$VERSION"
        return
    fi
    
    if [ -d "/Applications/Microsoft Excel.app" ]; then
        VERSION=$(defaults read "/Applications/Microsoft Excel.app/Contents/Info" CFBundleShortVersionString 2>/dev/null)
        echo "$VERSION"
        return
    fi
    
    echo "Unknown"
}

# Function to check if Office 2024 is installed
is_office_2024() {
    VERSION=$(identify_office_version)
    if [[ $VERSION == 16.89* ]]; then
        return 0 # It's Office 2024
    else
        return 1 # It's not Office 2024
    fi
}

# Function to remove Office
remove_office() {
    echo -e "${BLUE}Iniciando processo de remoção do Office...${NC}"
    
    RESET_TOOL="Microsoft_Office_Reset_1.9.1.pkg"
    DOWNLOAD_URL="https://office-reset.com/download/Microsoft_Office_Reset_1.9.1.pkg"
    
    echo -e "${BLUE}Baixando ferramenta de reset do Office...${NC}"
    curl -L -o "/tmp/$RESET_TOOL" "$DOWNLOAD_URL"
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}Falha ao baixar a ferramenta de reset. Tente novamente.${NC}"
        return 1
    fi
    
    echo -e "${GREEN}Executando ferramenta de reset. Siga as instruções na tela.${NC}"
    open "/tmp/$RESET_TOOL"
    
    echo -e "${BLUE}Aguardando o término da remoção. Pressione Enter quando finalizar.${NC}"
    read -p ""
    
    echo -e "${GREEN}Remoção concluída.${NC}"
    return 0
}

# Function to download and install Office 2024
install_office_2024() {
    echo -e "${BLUE}Iniciando instalação do Office 2024...${NC}"
    
    OFFICE_INSTALLER="Microsoft_365_and_Office_16.89.24090815_HomeStudent_Installer.pkg"
    DOWNLOAD_URL="https://officecdn.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/Microsoft_365_and_Office_16.89.24090815_HomeStudent_Installer.pkg"
    
    echo -e "${BLUE}Baixando o instalador do Office 2024...${NC}"
    curl -L -o "/tmp/$OFFICE_INSTALLER" "$DOWNLOAD_URL"
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}Falha ao baixar o instalador do Office. Tente novamente.${NC}"
        return 1
    fi
    
    echo -e "${GREEN}Executando o instalador do Office 2024. Siga as instruções na tela.${NC}"
    open "/tmp/$OFFICE_INSTALLER"
    
    echo -e "${BLUE}Aguardando o término da instalação. Pressione Enter quando finalizar.${NC}"
    read -p ""
    
    echo -e "${GREEN}Instalação concluída.${NC}"
    return 0
}

# Function to activate Office 2024
activate_office_2024() {
    echo -e "${BLUE}Iniciando ativação do Office 2024...${NC}"
    
    VL_SERIALIZER="Microsoft_Office_LTSC_2024_VL_Serializer.pkg"
    DOWNLOAD_URL="https://delivery.activated.win/dbmassgrave/Microsoft_Office_LTSC_2024_VL_Serializer.pkg?t=Kz9nBaNh5NkDOBbwe9Kwsaw8A0O29Srt&P1=1745893472&P2=601&P3=2&P4=%2B7CaifSrJo16xgalb62Kjrwwo2oOTXF%2BwbxBpfu5jC0%3D"
    
    echo -e "${BLUE}Baixando ferramenta de ativação...${NC}"
    curl -L -o "/tmp/$VL_SERIALIZER" "$DOWNLOAD_URL"
    
    if [ $? -ne 0 ]; then
        echo -e "${RED}Falha ao baixar a ferramenta de ativação. Tente novamente.${NC}"
        return 1
    fi
    
    echo -e "${GREEN}Executando ferramenta de ativação. Siga as instruções na tela.${NC}"
    open "/tmp/$VL_SERIALIZER"
    
    echo -e "${BLUE}Aguardando o término da ativação. Pressione Enter quando finalizar.${NC}"
    read -p ""
    
    echo -e "${GREEN}Ativação concluída.${NC}"
    return 0
}

# Main function
main() {
    echo -e "${BLUE}=== Sistema de Gerenciamento do Microsoft Office 2024 ===${NC}"
    
    if check_office_installed; then
        OFFICE_VERSION=$(identify_office_version)
        echo -e "${GREEN}Office instalado. Versão: $OFFICE_VERSION${NC}"
        
        if ! is_office_2024; then
            echo -e "${RED}Versão instalada não é Office 2024. Recomendamos remover.${NC}"
            read -p "Deseja remover a instalação atual do Office? (s/n): " REMOVE_CHOICE
            
            if [[ $REMOVE_CHOICE == "s" || $REMOVE_CHOICE == "S" ]]; then
                if remove_office; then
                    echo -e "${GREEN}Office removido com sucesso. Procedendo com nova instalação.${NC}"
                    install_office_2024 && activate_office_2024
                fi
            else
                echo -e "${RED}Operação cancelada pelo usuário.${NC}"
            fi
        else
            read -p "Office 2024 instalado. Deseja ativar (a) ou remover (r)? (a/r): " CHOICE
            
            if [[ $CHOICE == "a" || $CHOICE == "A" ]]; then
                activate_office_2024
            elif [[ $CHOICE == "r" || $CHOICE == "R" ]]; then
                remove_office
            else
                echo -e "${RED}Opção inválida.${NC}"
            fi
        fi
    else
        echo -e "${BLUE}Office não instalado.${NC}"
        read -p "Deseja instalar o Office 2024? (s/n): " INSTALL_CHOICE
        
        if [[ $INSTALL_CHOICE == "s" || $INSTALL_CHOICE == "S" ]]; then
            install_office_2024 && activate_office_2024
        else
            echo -e "${RED}Instalação cancelada pelo usuário.${NC}"
        fi
    fi
    
    echo -e "${GREEN}Operação concluída.${NC}"
}

# Execute main function
main
