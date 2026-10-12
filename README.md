# 🛒 E-Commerce Web Scraper & Dynamic Price Tracker

A Python automation tool designed to scrape product data, dynamic prices, availability, and review ratings from e-commerce platforms, exporting the extracted data into a custom-styled Excel spreadsheet.

## 🚀 Features

- **Automated Data Extraction**: Scrapes book titles, prices, star ratings, and stock status in real-time.
- **Robust Parsing**: Uses BeautifulSoup to extract nested CSS classes and attributes.
- **Custom Excel Styling**: Automatically creates headers with corporate fills, aligned typography, and dynamic column auto-fitting using openpyxl.
- **Modular Structure**: Easily adaptable for scraping product catalogs across different e-commerce architectures.

## 🛠️ Tech Stack

- Python 3.14+
- Requests (HTTP client for fetching web contents)
- BeautifulSoup4 (HTML parsing and structural navigation)
- OpenPyXL (Excel spreadsheet creation and styling)

## 📁 Repository Structure

projeto-3-ecommerce-scraping/
├── .gitignore
├── README.md
├── ecommerce_scraper.py
└── Relatorio_Ecommerce_Livros.xlsx

## 📋 Prerequisites & Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/anthonygbarros/projetos-portfolio.git]

## ⚡ How to Run

Execute the Python script:
python ecommerce_scraper.py

Upon execution, the script will generate Relatorio_Ecommerce_Livros.xlsx containing all parsed product records.