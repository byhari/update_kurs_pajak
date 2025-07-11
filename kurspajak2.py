import streamlit as st
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import io

# Base URL for scraping
base_url = "https://fiskal.kemenkeu.go.id/informasi-publik/kurs-pajak"

# Function to scrape data
def scrape_data():
    data = []
    failed_weeks = []  # To track failed weeks
    today = datetime.today()
    start_date = today - timedelta(weeks=13)  # 3 months ago
    current_date = start_date
    total_weeks = (today - start_date).days // 7 + 1
    week_count = 0
    progress = st.progress(0)

    while current_date <= today:
        # Align the date to the correct week format
        week_start = current_date + timedelta(days=(2 - current_date.weekday()) % 7)
        week_end = week_start + timedelta(days=6)
        week_number, year = week_start.isocalendar()[1], week_start.isocalendar()[0]
        week_code = f"{week_number:02}{year}"  # Adjusted to ISO week format

        # Request data
        params = {"date": week_start.strftime('%Y-%m-%d')}
        response = requests.get(base_url, params=params)

        if response.status_code != 200:
            failed_weeks.append(week_code)
            current_date += timedelta(days=7)
            week_count += 1
            progress.progress(min(1.0, week_count / total_weeks))
            continue

        # Parse response
        soup = BeautifulSoup(response.content, 'html.parser')
        rows = soup.find_all('tr', class_='table-bordered')

        if not rows:
            st.warning(f"No currency data found for week {week_code}. Skipping...")
            current_date += timedelta(days=7)
            week_count += 1
            progress.progress(min(1.0, week_count / total_weeks))
            continue

        for row in rows:
            # Extract currency name
            currency_name_tag = row.find('span', class_='hidden-xs')
            if not currency_name_tag or 'USD' not in currency_name_tag.text:
                continue

            # Extract value
            value_div = row.find('div', class_='m-l-5')
            value_text = value_div.text.strip() if value_div else 'N/A'

            try:
                # Correct numeric formatting
                if value_text.upper() == 'N/A':
                    raise ValueError("Value is N/A")
                
                # Convert value to float (handle commas and periods)
                value = float(value_text.replace('.', '').replace(',', '.'))  # Remove thousands separator and convert to float
            except ValueError:
                st.warning(f"Invalid VALUE format for {week_code}: {value_text}")
                value = None  # Set to None for invalid values

            # Append valid data with custom column names
            if value is not None:
                data.append({
                    "START_DATE": week_start.strftime('%Y-%m-%d'),
                    "END_DATE": week_end.strftime('%Y-%m-%d'),
                    "WEEK_CODE": week_code,
                    "CURRENCY": 'USD',
                    "KURS_PAJAK": value
                })

        current_date += timedelta(days=7)
        week_count += 1
        progress.progress(min(1.0, week_count / total_weeks))

    if failed_weeks:
        st.warning(f"Failed to fetch data for {len(failed_weeks)} weeks.")
    else:
        st.success("Scraping complete with no errors.")

    return data

# Function to create downloadable Excel file
def create_excel_download_link(df, filename="kurs_pajak_records.xlsx"):
    """
    Creates a download link for an Excel file from a DataFrame
    """
    output = io.BytesIO()
    # Use openpyxl as the engine instead of xlsxwriter
    try:
        # Try with to_excel direct method first (simpler)
        df.to_excel(output, index=False)
    except Exception as e:
        st.error(f"Excel writing error: {e}")
        st.info("Trying alternative method...")
        try:
            # If that fails, try with explicit engine
            df.to_excel(output, engine='openpyxl', index=False)
        except Exception as e2:
            # Final fallback to CSV if Excel writing fails
            st.error(f"Excel writing error with engine: {e2}")
            st.warning("Falling back to CSV format")
            return df.to_csv(index=False).encode('utf-8'), "csv"
    
    output.seek(0)
    return output.getvalue(), "xlsx"

# Streamlit interface
st.title("Update Kurs Pajak")
st.markdown("<h3 style='font-size:18px; margin-top:-10px;'>for Tax Payment - byhari</h3>", unsafe_allow_html=True)

if st.button("Scrape Data"):
    st.write("Fetching data from the last 3 months ...")
    scraped_data = scrape_data()

    if scraped_data:
        st.write("Scraped data successfully.")
        # Create DataFrame and display preview
        df = pd.DataFrame(scraped_data)
        st.dataframe(df)
        
        # Create download button for Excel file
        excel_data, file_type = create_excel_download_link(df)
        
        # Allow user to download the file
        if file_type == "xlsx":
            st.download_button(
                label="Download Excel file",
                data=excel_data,
                file_name="kurs_pajak_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.download_button(
                label="Download CSV file (Excel format unavailable)",
                data=excel_data,
                file_name="kurs_pajak_records.csv",
                mime="text/csv"
            )
    else:
        st.warning("No data was scraped. Ensure the source website is accessible.")
