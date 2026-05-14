# 🎴 BOT Card Assets Repository

[![Update Assets](https://github.com/SploeCyber/BOT-Assets/actions/workflows/update_assets.yml/badge.svg)](https://github.com/SploeCyber/BOT-Assets/actions/workflows/update_assets.yml)
[![Merge Datasets](https://github.com/SploeCyber/BOT-Assets/actions/workflows/merge_datasets.yml/badge.svg)](https://github.com/SploeCyber/BOT-Assets/actions/workflows/merge_datasets.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Professional-grade card asset management system** for the BOT (Battle of Thailand) card game. Automated pipeline for extracting, optimizing, and distributing card data and images from Google Sheets.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Repository Structure](#repository-structure)
- [Quick Start](#quick-start)
- [For Developers](#for-developers)
  - [Using the Assets](#using-the-assets)
  - [Asset Map API](#asset-map-api)
  - [Dataset Schema](#dataset-schema)
- [For Contributors](#for-contributors)
- [Automation & CI/CD](#automation--cicd)
- [Performance](#performance)
- [License](#license)

---

## 🎯 Overview

This repository serves as the **single source of truth** for all BOT card game assets, including:

- 🖼️ **High-resolution card images** (PNG & optimized WebP/PNG formats)
- 📊 **Structured card data** (JSON datasets with card stats, effects, and metadata)
- 🔄 **Automated pipelines** for asset generation from Google Sheets
- ⚡ **Optimized delivery** with sparse checkout and incremental builds

All assets are automatically synchronized from the official BOT card database via GitHub Actions workflows.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔄 **Auto-Sync** | Daily automated updates from Google Sheets |
| ⚡ **Sparse Checkout** | Clone only what you need (~90% faster) |
| 🖼️ **Image Optimization** | WebP & optimized PNG formats for web performance |
| 📦 **Structured Data** | Clean JSON datasets for easy integration |
| 🔍 **Change Detection** | Hash-based deduplication, only updates changed sheets |
| 🛡️ **Production-Ready** | CI/CD with error handling, retries, and validation |
| 📊 **Merged Dataset** | Single consolidated JSON with all cards for easy querying |

---

## 📁 Repository Structure

```
BOT-Assets/
├── assets/                          # Card assets organized by set
│   ├── BT01 - Welcome ตลิ่งชัน/    # Booster Pack 01
│   │   ├── card-data.json          # Card data for this set
│   │   ├── BT01-001-UR.png        # High-res image
│   │   └── BT01-001-UR_optimized.png  # Optimized for web
│   ├── SD01 - ตัวตึงไกรลาส/       # Starter Deck 01
│   ├── KD01 อวตารลงทัณฑ์/         # Special Deck 01
│   ├── ...                         # Other sets
│   └── all-cards.json             # Merged dataset (all sets)
│
├── scripts/                        # Automation scripts
│   ├── download_sheets.py         # Google Sheets downloader
│   ├── generate_info_json.py      # XLSX to JSON converter
│   ├── optimize_images.py         # Image optimizer (PNG/WebP)
│   ├── merge_datasets.py          # Dataset consolidation
│   └── card_overrides.py          # Card data overrides
│
├── .github/workflows/             # CI/CD pipelines
│   ├── update_assets.yml         # Asset update workflow
│   └── merge_datasets.yml        # Dataset merge workflow
│
└── asset-map.json                 # Index of all available sets
```

### Set Naming Convention

| Prefix | Type | Example |
|--------|------|---------|
| **BT** | Booster Pack | BT01, BT02, BT03 |
| **SD** | Starter Deck | SD01, SD02, SD03 |
| **KD** | Kudson Deck | KD01, KD02, KD03 |
| **CC** | Community Collection | CC01, CC02 |
| **SL** | Selection | SL01 |
| **ODY** | Odenya Series | ODY1 |
| **PRE/PromoCard** | Promotional Cards | PRE0, PromoCard |

---

## 🚀 Quick Start

### 1. Clone the Repository (Optimized)

**Don't clone the entire repo!** Use sparse checkout to save time and bandwidth:

```bash
# Initialize sparse checkout
git clone --filter=blob:none --sparse https://github.com/SploeCyber/BOT-Assets.git
cd BOT-Assets

# Checkout only what you need
git sparse-checkout set scripts/ asset-map.json assets/BT01\ -\ Welcome\ ตลิ่งชัน/
```

### 2. Access Card Data

```bash
# View available sets
cat asset-map.json

# Load all card data (merged dataset)
cat assets/all-cards.json

# Access specific set data
cat assets/BT01\ -\ Welcome\ ตลิ่งชัน/card-data.json
```

### 3. Use in Your Application

```javascript
// Load asset map
const assetMap = await fetch('/path/to/asset-map.json').then(r => r.json());

// Load card data for a specific set
const set = assetMap.find(s => s.name.includes('BT01'));
const cards = await fetch(`/path/to/${set.path.replace('dataset.json', 'card-data.json')}`).then(r => r.json());

// Display a card with optimized image
const card = cards[0];
const imageUrl = `assets/BT01 - Welcome ตลิ่งชัน/${card.ImagePath.replace('.png', '_optimized.png')}`;
```

---

## 👨‍💻 For Developers

### Using the Assets

#### Recommended Workflow

1. **Start with `asset-map.json`** - Get list of all available card sets
2. **Load `card-data.json`** - Fetch card data for specific sets you need
3. **Request optimized images** - Use `*_optimized.png` files for faster loading
4. **Use `all-cards.json`** - For applications needing all cards in one file

#### Image Path Resolution

```javascript
// Given card.ImagePath = "BT01-001-UR.png"
// And card._source = "BT01 - Welcome ตลิ่งชัน"

const imagePath = `assets/${card._source}/${card.ImagePath}`;
// Result: "assets/BT01 - Welcome ตลิ่งชัน/BT01-001-UR.png"

// For web-optimized version:
const optimizedPath = imagePath.replace('.png', '_optimized.png');
```

### Asset Map API

**Endpoint:** `asset-map.json`

**Response Format:**
```json
[
  {
    "name": "BT01 - Welcome ตลิ่งชัน",
    "path": "assets/BT01 - Welcome ตลิ่งชัน/dataset.json"
  },
  // ... more sets
]
```

### Dataset Schema

Each card in `card-data.json` follows this structure:

```typescript
interface Card {
  ImagePath: string;      // Filename of the card image
  Name: string;           // Card name (Thai)
  Type: string;           // Card type (Avatar, Magic, Life, Construct, etc.)
  Cost?: number;          // Play cost
  Gem?: number;           // Gem value
  Power?: number;         // Attack power
  Symbol?: string;        // Symbol/faction
  Color?: string;         // Card color (hex code)
  Print: string;          // Print code (e.g., "BT01-001")
  Rare: string;           // Rarity (UR, SR, R, C, etc.)
  SubType?: string;       // Subtype if applicable
  Details?: {             // Card effects and additional info
    "Main Effect"?: string;
    [key: string]: any;
  };
  _source?: string;       // Source set name (only in all_cards.json)
  [key: string]: any;     // Additional set-specific fields
}
```

---

## 🤝 For Contributors

### Updating Card Data

Card data is managed through Google Sheets and automatically synced. **Do not manually edit `dataset.json` files.**

To request changes:
1. Contact the BOT development team to update the Google Sheet
2. Trigger a manual workflow run if urgent update needed
3. Or submit an issue describing the needed changes

### Running Scripts Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Check for sheet changes
python scripts/download_sheets.py check

# Download specific sheet
python scripts/download_sheets.py download-one 1

# Generate JSON from downloaded XLSX
python scripts/generate_info_json.py

# Optimize images
python scripts/optimize_images.py "BT01 - Welcome ตลิ่งชัน"

# Merge all datasets
python scripts/merge_datasets.py
```

---

## 🔄 Automation & CI/CD

### Workflows

| Workflow | Trigger | Purpose |
|----------|---------|---------|
| **Update Assets** | Manual / Schedule | Download & process card data from Google Sheets |
| **Merge Datasets** | Push to main / Daily | Consolidate all sets into single `all_cards.json` |

### Update Assets Workflow

```
1. Check for changes (hash comparison) ⚡
   └─ If no changes → Skip build (saves resources)
   
2. For each changed sheet:
   ├─ Download XLSX
   ├─ Extract card images
   ├─ Generate dataset.json
   ├─ Optimize images
   ├─ Commit & push changes
   
3. Update hash tracking
```

**Features:**
- ✅ Sparse checkout for fast clones
- ✅ Hash-based change detection
- ✅ Automatic retry on failures
- ✅ Per-sheet atomic commits
- ✅ Progress tracking & summaries

---

## ⚡ Performance

### Optimizations Applied

| Optimization | Impact |
|--------------|--------|
| **Sparse Checkout** | ~90% faster clone times |
| **Hash-based Detection** | Only processes changed sheets |
| **Image Optimization** | 50-70% smaller file sizes |
| **Parallel Processing** | 4x faster image processing |
| **Dependency Caching** | Faster CI/CD runs |
| **Incremental Builds** | Minimal work per update |

### Storage Metrics

- **Full PNG Images**: ~500KB average per card
- **Optimized PNG**: ~200KB average per card (60% reduction)
- **WebP Format**: ~150KB average per card (70% reduction) 💡

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 📞 Support

- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/SploeCyber/BOT-Assets/issues)
- 💬 **Questions**: [Discussions](https://github.com/SploeCyber/BOT-Assets/discussions)

---

<div align="center">

**Made with ❤️ by the Battle of Talingchan Community**

</div>
