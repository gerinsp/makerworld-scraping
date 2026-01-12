
## Requirements

- Python 3.14
- pip
- Virtual environment

## Setup Environment

### 1. Clone repository
```bash
git clone https://github.com/gerinsp/makerworld-scraping.git
cd makerworld-scraping
```

### 2. Create Virtual Environment
```bash
python3 -m venv venv
```

### 3. Enable Virtual Environment
```bash
source venv/bin/activate
```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

### 5. Install ffmpeg
```bash
brew install ffmpeg
```

### 6. Install Chromium
```bash
python -m playwright install chromium
```

### Example Usage

```bash
python app.py \
  -k "Mechanical Artillery moving turret + recoil shoot" \
  -m 1 \
  --template "Shopee_mass_upload_2026-01-12_basic_template.xlsx" \
  -o output/shopee_ready.xlsx \
  --desc-template desc_template.txt \
  --category-id 101967 \
  --brand "No Brand" \
  --price 45000 \
  --stock 20 \
  --weight-kg 0.2 \
  --dims 12,12,4 \
  --sku-prefix MW-CABLE-WINDER \
  --meta-out output/makerworld_meta.csv \
  --download-dir output/downloads \
  --allow-gif
