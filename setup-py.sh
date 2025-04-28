#!/bin/bash

# Set color variables for better output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

show_help() {
    echo -e "${YELLOW}Python Environment Setup Script${NC}"
    echo "Usage: $0 [option]"
    echo "Options:"
    echo "  -h, --help       Show this help message"
    echo "  -c, --create        Create new virtual environment"
    echo "  -u, --update     Update requirements.txt"
    echo "  -i, --install    Install dependencies from requirements.txt"
    echo "  -n, --new        Setup a virtual environment and install dependencies with requirements.txt"
    echo ""
    echo "Examples:"
    echo "  $0 -n      # Create new virtual environment"
    echo "  $0 -a      # Full setup (venv + install)"
}

update_requirements() {
    printf "${GREEN}=== Updating Requirements.txt ===${NC}\n"
    if ! py -m pip freeze > requirements.txt; then
        printf "${RED}Error: Failed to update requirements.txt${NC}\n" >&2
        exit 1
    fi
    printf "${BLUE}Requirements.txt updated successfully${NC}\n"
}

create_new_venv() {
    if [[ -d "venv" ]]; then
        printf "${YELLOW}Virtual environment 'venv' already exists${NC}\n"
        return
    fi
    
    printf "${GREEN}=== Creating virtual environment => 'venv' ===${NC}\n"
    if ! py -m venv venv; then
        printf "${RED}Error: Failed to create virtual environment${NC}\n" >&2
        exit 1
    fi
    printf "${BLUE}Virtual environment created successfully${NC}\n"
}

install_dependencies() {
    if [[ -f "requirements.txt" ]]; then
        printf "${GREEN}=== Installing dependencies from Requirements.txt ===${NC}\n"
        if ! py -m pip install -r requirements.txt; then
            printf "${RED}Error: Failed to install dependencies${NC}\n" >&2
            exit 1
        fi
        printf "${BLUE}Dependencies installed successfully${NC}\n"
    else
        printf "${RED}=== Requirements.txt not found ===${NC}\n" >&2
        exit 1
    fi
}

setup_with_dependencies() {
    create_new_venv
    install_dependencies
}

case "$1" in
    -h|--help|help)
        show_help
        ;;
    -c|--create)
        create_new_venv
        ;;
    -u|--update)
        update_requirements
        ;;
    -i|--install)
        install_dependencies
        ;;
    -n|--new)
        setup_with_dependencies
        ;;
    *)
        printf "${RED}Error: Invalid option${NC}\n" >&2
        show_help
        exit 1
        ;;
esac

exit 0