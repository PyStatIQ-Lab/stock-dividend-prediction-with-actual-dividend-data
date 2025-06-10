import yfinance as yf
import pandas as pd
import os
from openpyxl import load_workbook
import streamlit as st

# Configure paths
STOCKS_FILE_PATH = 'stocks.xlsx'
ACTUAL_DIVIDEND_FILE = 'actual_stock-dividend.csv'

def safe_get(data, key, default="N/A"):
    """Helper function to safely get values from financial statements"""
    try:
        return data.loc[key].iloc[0] if key in data.index else default
    except Exception:
        return default

def get_financial_data(ticker):
    yf_ticker = ticker + ".NS"
    stock = yf.Ticker(yf_ticker)
    result = {'Ticker': ticker}
    
    # Fetch financial data with better error handling
    try:
        financials = {
            'income': stock.financials,
            'balance': stock.balance_sheet,
            'cashflow': stock.cashflow,
            'dividends': stock.dividends,
            'info': stock.info
        }
    except Exception as e:
        st.error(f"Error fetching data for {ticker}: {str(e)}")
        return None

    # Get basic price data
    try:
        latest_close = stock.history(period="1d")['Close'].iloc[-1]
    except Exception:
        latest_close = "N/A"
    result['Latest Close Price'] = latest_close

    # Financial metrics with proper fallbacks
    result.update({
        'Net Income': safe_get(financials['income'], 'Net Income'),
        'Operating Income': safe_get(financials['income'], 'Operating Income') or 
                         safe_get(financials['income'], 'EBIT'),
        'EPS': safe_get(financials['income'], 'Basic EPS') or
              (safe_get(financials['income'], 'Net Income') / 
               financials['info'].get('sharesOutstanding', 1) 
               if isinstance(safe_get(financials['income'], 'Net Income'), (int, float)) else "N/A"),
        'Revenue Growth': financials['income'].loc['Total Revenue'].pct_change().iloc[-1] 
                         if 'Total Revenue' in financials['income'].index else "N/A",
        'Retained Earnings': safe_get(financials['balance'], 'Retained Earnings'),
        'Cash Reserves': safe_get(financials['balance'], 'Cash And Cash Equivalents') or
                        safe_get(financials['balance'], 'Cash'),
        'Debt-to-Equity': (safe_get(financials['balance'], 'Total Debt') / 
                         safe_get(financials['balance'], 'Total Stockholder Equity')
                         if all(k in financials['balance'].index for k in ['Total Debt', 'Total Stockholder Equity']) else "N/A",
        'Dividend Yield': financials['info'].get('dividendYield', "N/A"),
        'Free Cash Flow': safe_get(financials['cashflow'], 'Free Cash Flow')
    })

    # Dividend analysis with better handling
    dividends = financials['dividends']
    if not isinstance(dividends, pd.Series) or dividends.empty:
        dividend_data = {
            'Dividend Growth Rate': "N/A",
            'Next Dividend Date': "N/A",
            'Predicted Dividend Amount': "N/A",
            'Dividend Percentage': "N/A",
            'Past Dividends': []
        }
    else:
        try:
            last_div = dividends.iloc[-1]
            div_growth = dividends.pct_change().mean()
            past_divs = dividends.tail(5).tolist()
            
            # Predict next dividend date
            if len(dividends) > 1:
                avg_period = (dividends.index[-1] - dividends.index[-2]).days
                next_date = dividends.index[-1] + pd.Timedelta(days=avg_period)
                next_date_str = next_date.strftime('%Y-%m-%d')
            else:
                next_date_str = "N/A"
            
            div_pct = (last_div/latest_close)*100 if isinstance(latest_close, (int, float)) else "N/A"
            
            dividend_data = {
                'Dividend Growth Rate': div_growth,
                'Next Dividend Date': next_date_str,
                'Predicted Dividend Amount': last_div,
                'Dividend Percentage': div_pct,
                'Past Dividends': past_divs
            }
        except Exception as e:
            st.error(f"Dividend calculation error for {ticker}: {str(e)}")
            dividend_data = {
                'Dividend Growth Rate': "N/A",
                'Next Dividend Date': "N/A",
                'Predicted Dividend Amount': "N/A",
                'Dividend Percentage': "N/A",
                'Past Dividends': []
            }
    
    result.update(dividend_data)

    # Compare with actual dividends
    try:
        actual_df = pd.read_csv(ACTUAL_DIVIDEND_FILE)
        actual_data = actual_df[actual_df['Symbol'].str.upper() == yf_ticker.upper()]
        
        if not actual_data.empty:
            latest_actual = actual_data.iloc[0]
            actual_div = float(latest_actual['Dividened Per Share'])
            actual_date = latest_actual['Ex Date']
            
            result.update({
                'Actual Dividend': actual_div,
                'Actual Ex Date': actual_date
            })
            
            # Calculate prediction accuracy if we have both values
            if (isinstance(result['Predicted Dividend Amount'], (int, float)) and 
                actual_div > 0):
                error = abs(result['Predicted Dividend Amount'] - actual_div)
                accuracy = max(0, 100 - (error/actual_div)*100)
                
                result.update({
                    'Dividend Prediction Error': error,
                    'Prediction Accuracy (%)': accuracy
                })
            else:
                result.update({
                    'Dividend Prediction Error': "N/A",
                    'Prediction Accuracy (%)': "N/A"
                })
        else:
            result.update({
                'Actual Dividend': "N/A",
                'Actual Ex Date': "N/A",
                'Dividend Prediction Error': "N/A",
                'Prediction Accuracy (%)': "N/A"
            })
    except Exception as e:
        st.error(f"Actual dividend error for {ticker}: {str(e)}")
        result.update({
            'Actual Dividend': "N/A",
            'Actual Ex Date': "N/A",
            'Dividend Prediction Error': "N/A",
            'Prediction Accuracy (%)': "N/A"
        })

    return result

# Function to save results to an Excel file
def save_to_excel(results, filename="dividend_predictions.xlsx"):
    try:
        results_df = pd.DataFrame(results)
        
        # Clean up the Past Dividends column for Excel
        if 'Past Dividends' in results_df.columns:
            results_df['Past Dividends'] = results_df['Past Dividends'].apply(
                lambda x: ', '.join(map(str, x)) if isinstance(x, list) else x
            )
        
        if os.path.exists(filename):
            book = load_workbook(filename)
            writer = pd.ExcelWriter(filename, engine='openpyxl')
            writer.book = book
            # Clear existing sheet if it exists
            if 'Dividend Predictions' in book.sheetnames:
                std = book.get_sheet_by_name('Dividend Predictions')
                book.remove_sheet(std)
            results_df.to_excel(writer, index=False, sheet_name='Dividend Predictions')
            writer.save()
        else:
            results_df.to_excel(filename, index=False)
        st.success(f"Results saved to {filename}")
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")

# Streamlit App
st.set_page_config(page_title="Stock Dividend Predictions", layout="wide")

# Display Header Logo
st.markdown("""
    <style>
        .header-logo {
            display: block;
            margin-left: auto;
            margin-right: auto;
            width: 25%;
        }
        /* Hide GitHub icons and fork button */
        .css-1v0mbdj { 
            display: none !important;
        }
        .css-1b22hs3 {
            display: none !important;
        }
        /* Hide Streamlit footer elements */
        footer { 
            display: none !important; 
        }
        /* Hide the GitHub repository button */
        .css-1r6ntm8 { 
            display: none !important;
        }
        /* Style the dataframe */
        .dataframe {
            width: 100%;
        }
        /* Style the progress bar */
        .stProgress > div > div > div > div {
            background-color: #4CAF50;
        }
    </style>
    <img class="header-logo" src="https://pystatiq.com/images/pystatIQ_logo.png" alt="Header Logo">
""", unsafe_allow_html=True)

st.title('Stock Dividend Prediction and Financial Analysis')

# Read the stock symbols from the local stocks.xlsx file
if os.path.exists(STOCKS_FILE_PATH):
    symbols_df = pd.read_excel(STOCKS_FILE_PATH)

    # Check if the 'Symbol' column exists
    if 'Symbol' not in symbols_df.columns:
        st.error("The file must contain a 'Symbol' column with stock tickers.")
    else:
        # Let the user select stocks from the file
        stock_options = symbols_df['Symbol'].tolist()
        selected_stocks = st.multiselect("Select Stock Symbols", stock_options)

        # Button to start the data fetching process
        if st.button('Fetch Financial Data') and selected_stocks:
            all_results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_stocks = len(selected_stocks)
            
            for i, ticker in enumerate(selected_stocks):
                progress = (i + 1) / total_stocks
                progress_bar.progress(progress)
                status_text.text(f"Processing {ticker} ({i+1}/{total_stocks})...")
                
                result = get_financial_data(ticker)
                if result is not None:
                    all_results.append(result)
            
            progress_bar.empty()
            status_text.empty()
            
            if all_results:
                st.subheader("Financial Data Results")
                results_df = pd.DataFrame(all_results)
                
                # Calculate accuracy metrics
                valid_errors = results_df[
                    (results_df['Dividend Prediction Error'] != "N/A") & 
                    (results_df['Actual Dividend'] != "N/A")
                ]
                
                if not valid_errors.empty:
                    st.write(f"### Prediction Accuracy Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Stocks with valid predictions", len(valid_errors))
                    with col2:
                        st.metric("Mean Absolute Error", f"{valid_errors['Dividend Prediction Error'].mean():.2f}")
                    with col3:
                        st.metric("Average Accuracy", f"{valid_errors['Prediction Accuracy (%)'].mean():.2f}%")
                
                # Display the full results
                st.dataframe(results_df)
                
                # Button to save the results to Excel
                if st.button('Save Results to Excel'):
                    save_to_excel(all_results)
                    st.balloons()

else:
    st.error(f"{STOCKS_FILE_PATH} not found. Please ensure the file exists.")

# Display Footer
st.markdown("""
    <div style="text-align: center; font-size: 14px; margin-top: 30px;">
        <p><strong>App Code:</strong> Stock-Dividend-Prediction-Jan-2025</p>
        <p>To get access to the stocks file to upload, please Email us at <a href="mailto:support@pystatiq.com">support@pystatiq.com</a>.</p>
        <p>Don't forget to add the Application code.</p>
        <p><strong>README:</strong> <a href="https://pystatiq-lab.gitbook.io/docs/python-apps/stock-dividend-predictions" target="_blank">https://pystatiq-lab.gitbook.io/docs/python-apps/stock-dividend-predictions</a></p>
    </div>
""", unsafe_allow_html=True)

# Display Footer Logo
st.markdown(f"""
    <style>
        .footer-logo {{
            display: block;
            margin-left: auto;
            margin-right: auto;
            width: 90px;
            padding-top: 30px;
        }}
    </style>
    <img class="footer-logo" src="https://predictram.com/images/logo.png" alt="Footer Logo">
""", unsafe_allow_html=True)
