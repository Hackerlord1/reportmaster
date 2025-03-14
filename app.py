from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Company configurations
COMPANIES = {
    'Canon': {
        'fsr_sales_targets': {
            'Jane Gichohi': 2000000,
            'Glayds  Bundotich': 1200000,
            'Beatrice Chege': 2600000,
            'Mirriam Makau': 1300000,
            'Victor Njogu': 1400000,
        },
        'unit_targets': {
            'Jane Gichohi': 'Retail',
            'Glayds  Bundotich': 'Retail',
            'Beatrice Chege': 'Retail/Marble',
            'Mirriam Makau': 'Retail',
            'Victor Njogu': 'Retail',
        },
        'aer_fsr_targets': {
            'Jane Gichohi': 228,
            'Glayds  Bundotich': 238,
            'Beatrice Chege': 180,
            'Mirriam Makau': 210,
            'Victor Njogu': 200,
        },
        'brand_targets': {
            'GOODNIGHT INSECT KILLER': {
                'Jane Gichohi': 15,
                'Glayds  Bundotich': 15,
                'Beatrice Chege': 12,
                'Mirriam Makau': 12,
                'Victor Njogu': 12,
            },
            'AER POWER POCKET': {
                'Jane Gichohi': 25,
                'Glayds  Bundotich': 20,
                'Beatrice Chege': 25,
                'Mirriam Makau': 20,
                'Victor Njogu': 25,
            },
            'FGWHHMG0N01': {
                'Jane Gichohi': 30,
                'Glayds  Bundotich': 25,
                'Beatrice Chege': 50,
                'Mirriam Makau': 25,
                'Victor Njogu': 40,
            },
            'FGWHHMG0N02': {
                'Jane Gichohi': 30,
                'Glayds  Bundotich': 30,
                'Beatrice Chege': 50,
                'Mirriam Makau': 30,
                'Victor Njogu': 40,
            },
            'FGWHTRMG0003': {  # New SKU
                'Jane Gichohi': 30,
                'Glayds  Bundotich': 30,
                'Beatrice Chege': 50,
                'Mirriam Makau': 30,
                'Victor Njogu': 40,
            },
        },
    },
    'Jumra': {
        'fsr_sales_targets': {
            'Jedidah Kemunto': 2000000,
            'Ochieng Charles': 1300000,
            'Lenah Cheloti': 1300000,
            'Moses  Ngugi': 3000000,
        },
        'unit_targets': {
            'Jedidah Kemunto': 'Retail',
            'Ochieng Charles': 'Retail',
            'Lenah Cheloti': 'Retail',
            'Moses  Ngugi': 'Retail',
        },
        'aer_fsr_targets': {
            'Jedidah Kemunto': 500,
            'Ochieng Charles': 500,
            'Lenah Cheloti': 500,
            'Moses  Ngugi': 500,
        },
        'sub_companies': {
            'Kimberly Clark': {
                'brands': ['Huggies', 'Kotex'],
                'fsr_targets': {
                    'Jedidah Kemunto': 1200000,
                    'Ochieng Charles': 800000,
                    'Lenah Cheloti': 700000,
                    'Moses  Ngugi': 1300000,
                },
                'eco_targets': {
                    'Jedidah Kemunto': 250,
                    'Ochieng Charles': 250,
                    'Lenah Cheloti': 250,
                    'Moses  Ngugi': 250,
                },
                'brand_targets': {
                    'Huggies': {
                        'Jedidah Kemunto': 100,
                        'Ochieng Charles': 100,
                        'Lenah Cheloti': 100,
                        'Moses  Ngugi': 100,
                    },
                    'Kotex': {
                        'Jedidah Kemunto': 150,
                        'Ochieng Charles': 150,
                        'Lenah Cheloti': 150,
                        'Moses  Ngugi': 150,
                    },
                },
            },
            'Dabur': {
                'brands': ['Amla', 'Hobby', 'Dabur', 'Vatika', 'Fem', 'ORS'],
                'fsr_targets': {
                    'Jedidah Kemunto': 800000,
                    'Ochieng Charles':900000,
                    'Lenah Cheloti': 800000,
                    'Moses  Ngugi': 1700000,
                },
                'eco_targets': {
                    'Jedidah Kemunto': 250,
                    'Ochieng Charles': 250,
                    'Lenah Cheloti': 250,
                    'Moses  Ngugi': 250,
                },
                'brand_targets': {
                    'Amla': {
                        'Jedidah Kemunto': 10,
                        'Ochieng Charles': 15,
                        'Lenah Cheloti': 8,
                        'Moses  Ngugi': 20,
                    },
                    'Hobby': {
                        'Jedidah Kemunto': 20,
                        'Ochieng Charles': 25,
                        'Lenah Cheloti': 18,
                        'Moses  Ngugi': 30,
                    },
                    'Dabur': {
                        'Jedidah Kemunto': 30,
                        'Ochieng Charles': 35,
                        'Lenah Cheloti': 28,
                        'Moses  Ngugi': 30,
                    },
                    'Vatika': {
                        'Jedidah Kemunto': 15,
                        'Ochieng Charles': 20,
                        'Lenah Cheloti': 12,
                        'Moses  Ngugi': 30,
                    },
                    'Fem': {
                        'Jedidah Kemunto': 25,
                        'Ochieng Charles': 30,
                        'Lenah Cheloti': 22,
                        'Moses  Ngugi': 25,
                    },
                    'ORS': {
                        'Jedidah Kemunto': 10,
                        'Ochieng Charles': 12,
                        'Lenah Cheloti': 8,
                        'Moses  Ngugi': 10,
                    },
                },
            },
        },
    },
}



def add_totals_row(df):
    """
    Add a totals row to the DataFrame and rename it to "KD Totals".
    - Sum numeric columns (e.g., Sales, ECO Actual).
    - Average percentage columns (e.g., % Sales, % ECO).
    """
    totals = {}
    for column in df.columns:
        if pd.api.types.is_numeric_dtype(df[column]):
            if column.startswith('%'):  # Check if the column is a percentage
                totals[column] = df[column].mean()  # Calculate average for percentage columns
            else:
                totals[column] = df[column].sum()  # Sum for other numeric columns
        elif column == "FSR":
            totals[column] = "KD Totals"  # Rename the totals row
        else:
            totals[column] = ""  # Non-numeric columns (e.g., Unit)
    totals_row = pd.DataFrame(totals, index=["KD Totals"])
    return pd.concat([df, totals_row], ignore_index=True)


def process_canon(file_path):
    try:
        # Read the uploaded file
        df = pd.read_excel(file_path)

        # Normalize SKU_Code and Brand columns
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        # Get company-specific configurations
        company_config = COMPANIES['Canon']

        # --- GENERAL SALES REPORT ---
        sales_report = df.groupby('FSR', as_index=False).agg({
            'Amount': 'sum',  # Total Sales
            'Customer': pd.Series.nunique  # ECO (unique customers)
        })
        sales_report.rename(columns={
            'Amount': 'Sales Actual',
            'Customer': 'ECO Actual'
        }, inplace=True)

        # Add Target, Unit, ECO Target, and other required columns
        sales_report['Sales Target'] = sales_report['FSR'].map(company_config['fsr_sales_targets']).fillna(0)
        sales_report['Unit'] = sales_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
        sales_report['ECO Target'] = sales_report['FSR'].map(company_config['aer_fsr_targets']).fillna(0)

        # Calculate Balance and % Performance for Sales and ECO
        sales_report['Sales Balance'] = sales_report['Sales Actual'] - sales_report['Sales Target']
        sales_report['% Sales'] = (sales_report['Sales Actual'] / sales_report['Sales Target']) * 100
        sales_report['ECO Balance'] = sales_report['ECO Actual'] - sales_report['ECO Target']
        sales_report['% ECO'] = (sales_report['ECO Actual'] / sales_report['ECO Target']) * 100

        # Reorder columns to match the desired structure
        sales_report = sales_report[[
            'FSR', 'Unit', 'Sales Target', 'Sales Actual', 'Sales Balance', '% Sales',
            'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
        ]]

        # Add a totals row to the sales report
        totals = {
            'FSR': 'KD Totals',
            'Unit': '',  # Empty for the totals row
            'Sales Target': sales_report['Sales Target'].sum(),
            'Sales Actual': sales_report['Sales Actual'].sum(),
            'Sales Balance': sales_report['Sales Balance'].sum(),
            '% Sales': sales_report['% Sales'].mean(),  # Average for percentage
            'ECO Target': sales_report['ECO Target'].sum(),
            'ECO Actual': sales_report['ECO Actual'].sum(),
            'ECO Balance': sales_report['ECO Balance'].sum(),
            '% ECO': sales_report['% ECO'].mean()  # Average for percentage
        }
        totals_row = pd.DataFrame(totals, index=["KD Totals"])
        sales_report = pd.concat([sales_report, totals_row], ignore_index=True)

        # Format amounts and percentages
        numeric_columns = ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']
        for col in numeric_columns:
            if pd.api.types.is_numeric_dtype(sales_report[col]):
                sales_report[col] = sales_report[col].apply(lambda x: f"{x:,.2f}")
        percentage_columns = ['% Sales', '% ECO']
        for col in percentage_columns:
            if pd.api.types.is_numeric_dtype(sales_report[col]):
                sales_report[col] = sales_report[col].apply(lambda x: f"{x:.2f}%")

        # Save the sales report
        sales_report_path = os.path.join(UPLOAD_FOLDER, 'Canon_sales_report.xlsx')
        sales_report.to_excel(sales_report_path, index=False)

        # --- BRAND-SPECIFIC ECO REPORTS ---
        brand_reports = {}
        brand_targets = company_config.get('brand_targets', {})  # Get brand targets if they exist

        for brand in df['Brand'].unique():  # Iterate over unique brands in the uploaded file
            brand_df = df[df['Brand'] == brand]
            if not brand_df.empty:
                eco_report = brand_df.groupby('FSR', as_index=False).agg({
                    'Customer': pd.Series.nunique  # ECO Actual
                })
                eco_report.rename(columns={
                    'Customer': 'ECO Actual'
                }, inplace=True)

                # Add ECO Target (if the brand exists in brand_targets, otherwise use 0)
                eco_report['ECO Target'] = eco_report['FSR'].map(brand_targets.get(brand, {})).fillna(0)

                # Calculate ECO Balance and % ECO
                eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
                eco_report['% ECO'] = (eco_report['ECO Actual'] / eco_report['ECO Target']) * 100

                # Add a totals row
                totals = {
                    'FSR': 'KD Totals',
                    'ECO Actual': eco_report['ECO Actual'].sum(),
                    'ECO Target': eco_report['ECO Target'].sum(),
                    'ECO Balance': eco_report['ECO Balance'].sum(),
                    '% ECO': eco_report['% ECO'].mean()  # Average for percentage
                }
                totals_row = pd.DataFrame(totals, index=["KD Totals"])
                eco_report = pd.concat([eco_report, totals_row], ignore_index=True)

                # Format amounts and percentages
                numeric_columns = ['ECO Target', 'ECO Actual', 'ECO Balance']
                for col in numeric_columns:
                    if pd.api.types.is_numeric_dtype(eco_report[col]):
                        eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}")
                percentage_columns = ['% ECO']
                for col in percentage_columns:
                    if pd.api.types.is_numeric_dtype(eco_report[col]):
                        eco_report[col] = eco_report[col].apply(lambda x: f"{x:.2f}%")

                # Save the report
                eco_report_path = os.path.join(UPLOAD_FOLDER, f'Canon_{brand}_eco_report.xlsx')
                eco_report.to_excel(eco_report_path, index=False)
                brand_reports[brand] = (eco_report, eco_report_path)
            else:
                brand_reports[brand] = None

        # --- SKU-SPECIFIC ECO REPORTS ---
        sku_eco_reports = {}
        specific_skus = ['FGWHHMG0N01', 'FGWHHMG0N02', 'FGWHTRMG0003']

        for sku in specific_skus:
            sku_df = df[df['SKU_Code'] == sku]
            if not sku_df.empty:
                eco_report = sku_df.groupby('FSR', as_index=False).agg({
                    'Customer': pd.Series.nunique  # ECO Actual
                })
                eco_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)

                # Add ECO Target (from company_config)
                eco_report['ECO Target'] = eco_report['FSR'].map(company_config['brand_targets'].get(sku, {})).fillna(0)

                # Calculate ECO Balance and % ECO
                eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
                eco_report['ECO %'] = (eco_report['ECO Actual'] / eco_report['ECO Target']) * 100

                # Add a totals row
                totals = {
                    'FSR': 'KD Totals',
                    'ECO Actual': eco_report['ECO Actual'].sum(),
                    'ECO Target': eco_report['ECO Target'].sum(),
                    'ECO Balance': eco_report['ECO Balance'].sum(),
                    'ECO %': eco_report['ECO %'].mean()  # Average for percentage
                }
                totals_row = pd.DataFrame(totals, index=["KD Totals"])
                eco_report = pd.concat([eco_report, totals_row], ignore_index=True)

                # Format amounts and percentages
                numeric_columns = ['ECO Target', 'ECO Actual', 'ECO Balance']
                for col in numeric_columns:
                    if pd.api.types.is_numeric_dtype(eco_report[col]):
                        eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}")
                percentage_columns = ['ECO %']
                for col in percentage_columns:
                    if pd.api.types.is_numeric_dtype(eco_report[col]):
                        eco_report[col] = eco_report[col].apply(lambda x: f"{x:.2f}%")

                # Save the report
                eco_report_path = os.path.join(UPLOAD_FOLDER, f'{sku}_eco_report.xlsx')
                eco_report.to_excel(eco_report_path, index=False)
                sku_eco_reports[sku] = eco_report
            else:
                sku_eco_reports[sku] = None

        return (
            sales_report, sales_report_path,
            brand_reports, sku_eco_reports
        )

    except Exception as e:
        print(f"Error processing Canon file: {e}")
        raise ValueError(f"Error processing the Canon file: {e}")


def process_jumra(file_path):
    try:
        # Read the uploaded file
        df = pd.read_excel(file_path)

        # Normalize SKU_Code and Brand columns
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        # Get company-specific configurations
        company_config = COMPANIES['Jumra']

        # --- GENERAL SALES REPORT ---
        sales_report = df.groupby('FSR', as_index=False).agg({
            'Amount': 'sum',  # Total Sales
            'Customer': pd.Series.nunique  # ECO (unique customers)
        })
        sales_report.rename(columns={
            'Amount': 'Sales Actual',
            'Customer': 'ECO Actual'
        }, inplace=True)

        # Add Target, Unit, ECO Target, and other required columns
        sales_report['Sales Target'] = sales_report['FSR'].map(company_config['fsr_sales_targets']).fillna(0)
        sales_report['Unit'] = sales_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
        sales_report['ECO Target'] = sales_report['FSR'].map(company_config['aer_fsr_targets']).fillna(0)

        # Calculate Balance and % Performance for Sales and ECO
        sales_report['Sales Balance'] = sales_report['Sales Actual'] - sales_report['Sales Target']
        sales_report['% Sales'] = (sales_report['Sales Actual'] / sales_report['Sales Target']) * 100
        sales_report['ECO Balance'] = sales_report['ECO Actual'] - sales_report['ECO Target']
        sales_report['% ECO'] = (sales_report['ECO Actual'] / sales_report['ECO Target']) * 100

        # Reorder columns to match the desired structure
        sales_report = sales_report[[
            'FSR', 'Unit', 'Sales Target', 'Sales Actual', 'Sales Balance', '% Sales',
            'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
        ]]

        # Add a totals row to the sales report
        totals = {
            'FSR': 'KD Totals',
            'Unit': '',  # Empty for the totals row
            'Sales Target': sales_report['Sales Target'].sum(),
            'Sales Actual': sales_report['Sales Actual'].sum(),
            'Sales Balance': sales_report['Sales Balance'].sum(),
            '% Sales': sales_report['% Sales'].mean(),  # Average for percentage
            'ECO Target': sales_report['ECO Target'].sum(),
            'ECO Actual': sales_report['ECO Actual'].sum(),
            'ECO Balance': sales_report['ECO Balance'].sum(),
            '% ECO': sales_report['% ECO'].mean()  # Average for percentage
        }
        totals_row = pd.DataFrame(totals, index=["KD Totals"])
        sales_report = pd.concat([sales_report, totals_row], ignore_index=True)

        # Format amounts and percentages
        numeric_columns = ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']
        for col in numeric_columns:
            if pd.api.types.is_numeric_dtype(sales_report[col]):
                sales_report[col] = sales_report[col].apply(lambda x: f"{x:,.2f}")
        percentage_columns = ['% Sales', '% ECO']
        for col in percentage_columns:
            if pd.api.types.is_numeric_dtype(sales_report[col]):
                sales_report[col] = sales_report[col].apply(lambda x: f"{x:.2f}%")

        # Save the sales report
        sales_report_path = os.path.join(UPLOAD_FOLDER, 'Jumra_sales_report.xlsx')
        sales_report.to_excel(sales_report_path, index=False)

        # --- SUB-COMPANY REPORTS ---
        sub_company_reports = {}
        for sub_company, config in company_config['sub_companies'].items():
            sub_company_brands = config['brands']
            sub_company_targets = config['fsr_targets']
            sub_company_eco_targets = config.get('eco_targets', {})  # ECO Targets for the sub-company

            # Filter data for the sub-company's brands
            sub_company_df = df[df['Brand'].isin([b.upper() for b in sub_company_brands])]
            if not sub_company_df.empty:
                # Group by FSR and calculate total sales and ECO
                sub_company_report = sub_company_df.groupby('FSR', as_index=False).agg({
                    'Amount': 'sum',  # Actual Sales
                    'Customer': pd.Series.nunique  # ECO Actual
                })
                sub_company_report.rename(columns={
                    'Amount': 'Actual Sales',
                    'Customer': 'ECO Actual'
                }, inplace=True)

                # Add Sales Target, ECO Target, and Unit
                sub_company_report['Sales Target'] = sub_company_report['FSR'].map(sub_company_targets).fillna(0)
                sub_company_report['ECO Target'] = sub_company_report['FSR'].map(sub_company_eco_targets).fillna(0)
                sub_company_report['Unit'] = sub_company_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')

                # Calculate Sales Balance and % Sales
                sub_company_report['Sales Balance'] = sub_company_report['Actual Sales'] - sub_company_report['Sales Target']
                sub_company_report['% Sales'] = (sub_company_report['Actual Sales'] / sub_company_report['Sales Target']) * 100

                # Calculate ECO Balance and % ECO
                sub_company_report['ECO Balance'] = sub_company_report['ECO Actual'] - sub_company_report['ECO Target']
                sub_company_report['% ECO'] = (sub_company_report['ECO Actual'] / sub_company_report['ECO Target']) * 100

                # Reorder columns to match the desired structure
                sub_company_report = sub_company_report[[
                    'FSR', 'Unit', 'Sales Target', 'Actual Sales', 'Sales Balance', '% Sales',
                    'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
                ]]

                # Add a totals row to the sub-company report
                totals = {
                    'FSR': 'KD Totals',
                    'Unit': '',  # Empty for the totals row
                    'Sales Target': sub_company_report['Sales Target'].sum(),
                    'Actual Sales': sub_company_report['Actual Sales'].sum(),
                    'Sales Balance': sub_company_report['Sales Balance'].sum(),
                    '% Sales': sub_company_report['% Sales'].mean(),  # Average for percentage
                    'ECO Target': sub_company_report['ECO Target'].sum(),
                    'ECO Actual': sub_company_report['ECO Actual'].sum(),
                    'ECO Balance': sub_company_report['ECO Balance'].sum(),
                    '% ECO': sub_company_report['% ECO'].mean()  # Average for percentage
                }
                totals_row = pd.DataFrame(totals, index=["KD Totals"])
                sub_company_report = pd.concat([sub_company_report, totals_row], ignore_index=True)

                # Format amounts and percentages
                numeric_columns = ['Sales Target', 'Actual Sales', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']
                for col in numeric_columns:
                    if pd.api.types.is_numeric_dtype(sub_company_report[col]):
                        sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:,.2f}")
                percentage_columns = ['% Sales', '% ECO']
                for col in percentage_columns:
                    if pd.api.types.is_numeric_dtype(sub_company_report[col]):
                        sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:.2f}%")

                # Save the report
                sub_company_report_path = os.path.join(UPLOAD_FOLDER, f'Jumra_{sub_company}_report.xlsx')
                sub_company_report.to_excel(sub_company_report_path, index=False)

                # Add to sub_company_reports dictionary
                sub_company_reports[sub_company] = (sub_company_report, sub_company_report_path)
            else:
                sub_company_reports[sub_company] = (None, None)

        # --- ECO REPORTS FOR ALL BRANDS ---
        eco_reports = {}
        for sub_company, config in company_config['sub_companies'].items():
            brand_targets = config.get('brand_targets', {})
            for brand, targets in brand_targets.items():
                brand_df = df[df['Brand'] == brand.upper()]
                if not brand_df.empty:
                    eco_report = brand_df.groupby('FSR', as_index=False).agg({
                        'Customer': pd.Series.nunique  # ECO Actual
                    })
                    eco_report.rename(columns={
                        'Customer': 'ECO Actual'
                    }, inplace=True)

                    # Add ECO Target
                    eco_report['ECO Target'] = eco_report['FSR'].map(targets).fillna(0)

                    # Calculate ECO Balance and % ECO
                    eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
                    eco_report['% ECO'] = (eco_report['ECO Actual'] / eco_report['ECO Target']) * 100

                    # Add a totals row
                    totals = {
                        'FSR': 'KD Totals',
                        'ECO Actual': eco_report['ECO Actual'].sum(),
                        'ECO Target': eco_report['ECO Target'].sum(),
                        'ECO Balance': eco_report['ECO Balance'].sum(),
                        '% ECO': eco_report['% ECO'].mean()  # Average for percentage
                    }
                    totals_row = pd.DataFrame(totals, index=["KD Totals"])
                    eco_report = pd.concat([eco_report, totals_row], ignore_index=True)

                    # Format amounts and percentages
                    numeric_columns = ['ECO Target', 'ECO Actual', 'ECO Balance']
                    for col in numeric_columns:
                        if pd.api.types.is_numeric_dtype(eco_report[col]):
                            eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}")
                    percentage_columns = ['% ECO']
                    for col in percentage_columns:
                        if pd.api.types.is_numeric_dtype(eco_report[col]):
                            eco_report[col] = eco_report[col].apply(lambda x: f"{x:.2f}%")

                    # Save the report
                    eco_report_path = os.path.join(UPLOAD_FOLDER, f'{sub_company}_{brand}_eco_report.xlsx')
                    eco_report.to_excel(eco_report_path, index=False)
                    eco_reports[brand] = (eco_report, eco_report_path)
                else:
                    eco_reports[brand] = None

        return (
            sales_report, sales_report_path,
            sub_company_reports, eco_reports
        )

    except Exception as e:
        print(f"Error processing Jumra file: {e}")
        raise ValueError(f"Error processing the Jumra file: {e}")


def process_excel(file_path, company):
    if company == 'Canon':
        return process_canon(file_path)
    elif company == 'Jumra':
        return process_jumra(file_path)
    else:
        raise ValueError(f"Invalid company: {company}")


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded.'

        file = request.files['file']
        company = request.form.get('company')

        if file.filename == '':
            return 'No file selected.'

        if not company:
            return 'No company selected.'

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        try:
            result = process_excel(file_path, company)
            if result is None:
                return 'Error processing the file.'

            # Get today's date
            today_date = datetime.today().strftime('%Y-%m-%d')

            # Render the appropriate template based on the company
            if company == 'Canon':
                sales_report, sales_report_path, brand_reports, sku_eco_reports = result
                return render_template(
                    "canon_report.html",
                    sales_report=sales_report.to_html(classes="table table-striped table-bordered", index=False),
                    brand_reports=brand_reports,
                    sku_eco_reports=sku_eco_reports,
                    today_date=today_date
                )
            elif company == 'Jumra':
                sales_report, sales_report_path, sub_company_reports, eco_reports = result
                return render_template(
                    "jumra_report.html",
                    sales_report=sales_report.to_html(classes="table table-striped table-bordered", index=False),
                    sub_company_reports=sub_company_reports,
                    eco_reports=eco_reports,
                    today_date=today_date
                )
        except Exception as e:
            return f'Error: {e}'
        finally:
            if os.path.exists(file_path):
                os.remove(file_path)

    return '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Upload Sales Data</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
        <div class="container mt-5">
            <h1 class="text-center">Upload Sales Data</h1>
            <form method="post" enctype="multipart/form-data" class="mt-4">
                <div class="mb-3">
                    <label for="company" class="form-label">Select Company</label>
                    <select class="form-control" id="company" name="company" required>
                        <option value="">Choose a company</option>
                        <option value="Canon">Canon</option>
                        <option value="Jumra">Jumra</option>
                    </select>
                </div>
                <div class="mb-3">
                    <input type="file" class="form-control" name="file" accept=".xlsx" required>
                </div>
                <button type="submit" class="btn btn-primary">Upload and Generate Reports</button>
            </form>
        </div>
    </body>
    </html>
    '''


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)