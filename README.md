# ⚡ Sirocco I LP Loan Dashboard

A comprehensive Streamlit-based dashboard for managing and analyzing loan participation portfolios with life insurance integration.

## 🚀 Features

- **Loan Portfolio Management**: Track active, closed, and not-started loans
- **Life Insurance Portfolio Statistics**: Monitor face value, premiums, and management fees
- **Cash Flow Projections**: 6-month forward-looking cash flow analysis
- **Amortization Schedules**: Detailed payment tracking and analysis
- **Professional UI**: Dark theme with Sirocco branding
- **File Upload Support**: Excel and CSV file processing
- **Real-time Analytics**: Portfolio metrics and performance indicators

## 📊 Dashboard Sections

1. **Portfolio Summary**: Key metrics and performance indicators
2. **Active Loans**: Currently active loan portfolio
3. **Closed Loans**: Completed loan history
4. **Not Started Loans**: Upcoming loan commitments
5. **Cash Flow Projection**: 6-month payment forecasts
6. **Life Insurance Statistics**: Policy metrics and premium tracking
7. **Monthly Remittance Analysis**: Payment processing and reconciliation

## 🛠️ Local Development

### Prerequisites
- Python 3.11+
- pip package manager

### Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd sirocco-dashboard
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**:
   ```bash
   streamlit run streamlit_dashboard.py
   ```

5. **Access the dashboard**:
   Open your browser and go to `http://localhost:8501`

## 🌐 Heroku Deployment

### Prerequisites
- Heroku CLI installed
- Git repository initialized

### Deployment Steps

1. **Initialize Git repository** (if not already done):
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Sirocco dashboard"
   ```

2. **Login to Heroku**:
   ```bash
   heroku login
   ```

3. **Create Heroku app**:
   ```bash
   heroku create sirocco-loan-dashboard
   # Or let Heroku generate a name: heroku create
   ```

4. **Deploy to Heroku**:
   ```bash
   git push heroku main
   # Or if your branch is called master: git push heroku master
   ```

5. **Configure the app** (optional but recommended):
   ```bash
   # Set password for basic authentication
   heroku config:set PASSWORD=your-secure-password
   
   # Scale the dyno
   heroku ps:scale web=1
   ```

6. **Access your app**:
   ```bash
   heroku open
   ```

### Environment Variables

- `PASSWORD`: Set a secure password for basic authentication (default: "sirocco2024")

## 📁 File Structure

```
sirocco-dashboard/
├── streamlit_dashboard.py    # Main application file
├── requirements.txt          # Python dependencies
├── setup.sh                 # Heroku setup script
├── Procfile                 # Heroku process definition
├── runtime.txt              # Python version specification
├── .gitignore              # Git ignore rules
└── README.md               # This file
```

## 🔐 Security

The dashboard includes optional basic authentication:
- Default password: `sirocco2024`
- Can be customized via Heroku config vars
- Password is required to access the dashboard

## 📋 Expected Excel File Structure

The dashboard expects a Master Excel file with:

### Dashboard Sheet
- **E3**: As-of date
- **E39-E53**: Life insurance portfolio statistics
- **H39-H41**: Additional management metrics

### Loan Sheets (named with # prefix)
Each loan sheet should contain:

**Header Information** (either Format 1 or Format 2):
- **Format 1** (Data in column B):
  - A2 or B2: Borrower name
  - B3: Original loan amount
  - B4: Annual interest rate
  - B5: Loan period in months
  - B6: Payment amount
  - B7: Loan start date

- **Format 2** (Data in column C):
  - A2 or B2: Borrower name
  - C3: Original loan amount (when B3 contains label)
  - C4: Annual interest rate
  - C5: Loan period in months
  - C6: Payment amount
  - C7: Loan start date

**Amortization Schedule** (starting from row 11):
- A: Month
- B: Repayment number
- C: Opening balance
- D: Loan repayment
- E: Interest charged
- F: Capital repaid
- G: Closing balance
- J: Payment date
- K: Amount paid
- L: Notes

## 🎨 Customization

### Branding Colors
- Primary: `#FDB813` (Sirocco Yellow)
- Background: `#1a1a1a` (Dark)
- Secondary: `#2d2d2d` (Medium Dark)
- Accent: `#3d3d3d` (Light Dark)

### CSS Styling
The dashboard uses custom CSS for consistent branding. Modify the CSS section in `streamlit_dashboard.py` to customize the appearance.

## 🔧 Troubleshooting

### Common Issues

1. **Port already in use**:
   ```bash
   streamlit run streamlit_dashboard.py --server.port 8502
   ```

2. **Heroku deployment fails**:
   - Check that all files are committed to Git
   - Verify `Procfile` has no file extension
   - Ensure `requirements.txt` includes all dependencies

3. **Excel file not loading**:
   - Verify file format is `.xlsx`
   - Check that loan sheets start with `#`
   - Ensure Dashboard sheet exists

### Debug Mode
Enable debug information by checking "Show debug info" in the dashboard to see processing details.

## 📞 Support

For technical support or feature requests, please contact the development team.

## 📄 License

This project is proprietary software developed for Sirocco Partners.

---

**⚡ Powered by Sirocco Partners** 