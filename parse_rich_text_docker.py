#!/usr/bin/env python3
"""
Parse Excel cells using LibreOffice UNO running in a Docker container.
This is a client script that connects to LibreOffice running in Docker.

Requirements:
- Docker and docker-compose installed
- LibreOffice container running (see docker-compose.libreoffice.yml)

Usage:
    # Start the LibreOffice container
    docker-compose -f docker-compose.libreoffice.yml up -d
    
    # Run the parser
    python parse_rich_text_docker.py Book.xlsx
    
    # Stop the container when done
    docker-compose -f docker-compose.libreoffice.yml down
"""

import subprocess
import sys
import time
from pathlib import Path


def check_docker_installed() -> bool:
    """Check if Docker is installed and running."""
    try:
        result = subprocess.run(
            ["docker", "info"],
            capture_output=True,
            text=True,
            timeout=5
        )
        return result.returncode == 0
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return False


def check_container_running() -> bool:
    """Check if LibreOffice container is running."""
    try:
        result = subprocess.run(
            ["docker", "ps", "--filter", "name=libreoffice-uno", "--format", "{{.Names}}"],
            capture_output=True,
            text=True,
            timeout=5
        )
        return "libreoffice-uno" in result.stdout
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return False


def start_container() -> bool:
    """Start the LibreOffice Docker container."""
    print("Starting LibreOffice Docker container...")
    try:
        result = subprocess.run(
            ["docker-compose", "-f", "docker-compose.libreoffice.yml", "up", "-d"],
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            print(f"Error starting container: {result.stderr}")
            return False
        
        # Wait for container to be healthy
        print("Waiting for LibreOffice to be ready...", end="", flush=True)
        for i in range(30):
            time.sleep(1)
            print(".", end="", flush=True)
            
            # Check if container is healthy
            result = subprocess.run(
                ["docker", "inspect", "--format", "{{.State.Health.Status}}", "libreoffice-uno"],
                capture_output=True,
                text=True,
                timeout=5
            )
            
            if "healthy" in result.stdout:
                print(" Ready!")
                return True
        
        print(" Timeout!")
        return False
        
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        print(f"Error: {e}")
        return False


def run_parser_in_container(file_path: str) -> int:
    """
    Run the parser script inside the Docker container.
    
    Args:
        file_path: Path to Excel file (relative to current directory)
    
    Returns:
        Exit code from the parser script
    """
    # Make sure file exists
    if not Path(file_path).exists():
        print(f"Error: File not found: {file_path}")
        return 1
    
    print(f"\nParsing {file_path} using LibreOffice in Docker...")
    print("=" * 80)
    
    # Run the parser inside the container
    # The container has /app mounted to current directory
    cmd = [
        "docker", "exec", "-it", "libreoffice-uno",
        "python3", "/app/parse_rich_text_libreoffice.py",
        f"/app/{file_path}"
    ]
    
    try:
        result = subprocess.run(cmd)
        return result.returncode
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
        return 130


def main():
    """Main function."""
    # Check arguments
    if len(sys.argv) < 2:
        print("Usage: python parse_rich_text_docker.py <excel_file>")
        print("\nExample:")
        print("  python parse_rich_text_docker.py Book.xlsx")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    # Check Docker
    print("Checking Docker installation...")
    if not check_docker_installed():
        print("ERROR: Docker is not installed or not running")
        print("\nInstall Docker:")
        print("  macOS: https://docs.docker.com/desktop/install/mac-install/")
        print("  Linux: https://docs.docker.com/engine/install/")
        sys.exit(1)
    print("  Docker is installed and running")
    
    # Check if container is running
    print("\nChecking LibreOffice container...")
    if not check_container_running():
        print("  Container not running, starting it...")
        if not start_container():
            print("\nERROR: Failed to start LibreOffice container")
            print("\nTry manually:")
            print("  docker-compose -f docker-compose.libreoffice.yml up -d")
            sys.exit(1)
    else:
        print("  Container is already running")
    
    # Run the parser
    exit_code = run_parser_in_container(file_path)
    
    if exit_code == 0:
        print("\n" + "=" * 80)
        print("Parsing completed successfully!")
        print("\nTo stop the LibreOffice container:")
        print("  docker-compose -f docker-compose.libreoffice.yml down")
    
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
