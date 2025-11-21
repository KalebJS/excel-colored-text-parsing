# Excel Colored Text Parser

Parse Excel cells with colored text using LibreOffice UNO in Docker via Robocorp.

## Quick Start

```bash
./run_rcc_bot.sh
```

This will:
- Build the Docker image with LibreOffice and RCC
- Run the Robocorp bot to parse `Book.xlsx`
- Save results to `output/parsed_cells.txt`

## Requirements

- Docker
- Docker Compose

## How It Works

The bot uses:
- **LibreOffice UNO** to access rich text formatting in Excel cells
- **Robocorp RCC** to manage the Python environment and task execution
- **Docker** to provide a consistent Linux environment (works on Apple Silicon Macs)

## Files

- `run_rcc_bot.sh` - Main script to run the bot
- `robot.yaml` - Robocorp task configuration
- `tasks.py` - Main task entry point
- `tech_libreoffice.py` - LibreOffice UNO parsing logic
- `Dockerfile.rcc` - Docker image with LibreOffice and RCC
- `docker-compose.rcc.yml` - Docker Compose configuration

## Output

Results are saved to `output/parsed_cells.txt` with:
- RGB color values for each text segment
- Text content
- Color classification (black, red, blue, etc.)
