
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

### 5. Install Chromium
```bash
python -m playwright install chromium
```

### Example Usage

```bash
python3 app.py \
  -k "cable winder" \
  -m 1 \
  --template "Shopee_mass_upload_2025-10-18_100644.xlsx" \
  -o shopee_ready.xlsx \
  --category-id 120039 \
  --brand "No Brand" \
  --price 45000 \
  --stock 20 \
  --weight-kg 0.2 \
  --dims 12,12,4 \
  --sku-prefix MW-CABLE-WINDER
