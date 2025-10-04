# 💼 Artel Financial Forecast App

A Streamlit-based financial forecasting and reporting tool for **Artel Support LLC**.  
This app automates monthly revenue analysis from **SAP Excel exports**, helping finance teams quickly generate IFRS-ready reports, forecasts, and YoY comparisons — all **without coding**.

---

## 🚀 Features

- 📤 **Upload SAP Excel files** (current and previous periods)
- 🧾 **Automatic column translation** (Russian → English)
- 💰 **Call Center revenue** and **Admin cost forecast** input
- 🧮 **Forecast calculation**:
  - Revenue (After VAT)
  - G1 transport deductions
  - Net Revenue, VAT 12%, and Revenue After VAT
  - Forecasted monthly revenue (active days × daily avg)
- 📈 **YoY comparison**:
  - Projected revenue vs. last year’s actual
  - Daily revenue dynamics
  - Warranty structure (G1, G2, G3 shares)
- 📊 **Excel export** with multiple sheets:
  - Summary  
  - Reconciliation  
  - P&L Forecast  
  - Daily Revenue  
  - YoY Analysis  
  - Data Quality Checks

---

## 🧠 How It Works

1. Upload your **SAP Excel report (.xls/.xlsx)**  
2. Enter:
   - Call Center revenue (USD, VAT-included)
   - Admin cost forecast (USD, after VAT / net)
3. Choose current and previous periods (e.g., `2025-10` and `2024-10`)
4. Click **Run Forecast**
5. Download the **final Excel report** with all tables & analysis

---

## 🛠️ Tech Stack

- **Python 3.10+**
- **Streamlit** – UI & web app
- **Pandas** – data processing
- **OpenPyXL** – Excel export
- **NumPy** – calculations

---

## 📦 Installation (For Developers)

```bash
# Clone repo
git clone https://github.com/YOUR-USERNAME/artel-financial-forecast.git
cd artel-financial-forecast

# Install dependencies
pip install -r requirements.txt

# Run locally
streamlit run app.py
