from flask import Flask, render_template, request, send_file, session
from openpyxl.styles import Font, PatternFill, numbers
import pandas as pd
import os
from datetime import datetime
import io
import logging


# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


import os
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'uzumymw0')

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

COMPANIES = {
    'Canon': {
        'fsr_sales_targets': {
            'Jane Gichohi': 2000000,
            'Gladys Bundotich': 1200000,
            'Beatrice Chege': 2600000,
            'Mirriam Makau': 1300000,
            'Victor Njogu': 1400000,
        },
        'unit_targets': {
            'Jane Gichohi': 'Retail',
            'Gladys Bundotich': 'Retail',
            'Beatrice Chege': 'Retail/Marble',
            'Mirriam Makau': 'Retail',
            'Victor Njogu': 'Retail',
        },
        'aer_fsr_targets': {
            'Jane Gichohi': 228,
            'Gladys Bundotich': 238,
            'Beatrice Chege': 180,
            'Mirriam Makau': 210,
            'Victor Njogu': 200,
        },
        'brand_targets': {
            'GOODNIGHT INSECT KILLER': {
                'Jane Gichohi': 15,
                'Gladys Bundotich': 15,
                'Beatrice Chege': 12,
                'Mirriam Makau': 12,
                'Victor Njogu': 12,
            },
            'AER POWER POCKET': {
                'Jane Gichohi': 25,
                'Gladys Bundotich': 20,
                'Beatrice Chege': 25,
                'Mirriam Makau': 20,
                'Victor Njogu': 25,
            },
            'FGWHHMG0N01': {
                'Jane Gichohi': 30,
                'Gladys Bundotich': 25,
                'Beatrice Chege': 50,
                'Mirriam Makau': 25,
                'Victor Njogu': 40,
            },
            'FGWHHMG0N02': {
                'Jane Gichohi': 30,
                'Gladys Bundotich': 30,
                'Beatrice Chege': 50,
                'Mirriam Makau': 30,
                'Victor Njogu': 40,
            },
            'FGWHTRMG0003': {
                'Jane Gichohi': 30,
                'Gladys Bundotich': 30,
                'Beatrice Chege': 50,
                'Mirriam Makau': 30,
                'Victor Njogu': 40,
            },
        },
    },
    'Canon Eldoret': {
        'fsr_sales_targets': {
            'ANGELA AKIRAPA (Till 4545130)': 1740000,
            'BEATRICE INGABIRE(Till 4545130)': 2784000,
            'FANCY CHEPNGENO ': 928000,
            'IRENE MUNYOKA': 812000,
            'EZEKIEL NJOROGE': 1044000,
            'Walkin  Sales': 3132000.
        },
        'unit_targets': {
            'ANGELA AKIRAPA (Till 4545130)': 'Retail',
            'BEATRICE INGABIRE(Till 4545130)': 'Retail/Marble',
            'FANCY CHEPNGENO ': 'Retail',
            'IRENE MUNYOKA': 'Retail',
            'EZEKIEL NJOROGE': 'Ratail',
            'Walkin  Sales': 'Wholesale',
        },
        'aer_fsr_targets': {
            'ANGELA AKIRAPA (Till 4545130)': 150,
            'BEATRICE INGABIRE(Till 4545130)': 150,
            'FANCY CHEPNGENO ': 250,
            'IRENE MUNYOKA': 250,
            'EZEKIEL NJOROGE': 250,
            'Walkin  Sales': 23,
        },
        'brand_targets': {
            'GOODNIGHT INSECT KILLER': {
                'ANGELA AKIRAPA (Till 4545130)': 12,
                'BEATRICE INGABIRE(Till 4545130)': 12,
                'FANCY CHEPNGENO ': 12,
                'IRENE MUNYOKA': 12,
                'EZEKIEL NJOROGE': 15,
            },
            'AER POWER POCKET': {
                'ANGELA AKIRAPA (Till 4545130)': 15,
                'BEATRICE INGABIRE(Till 4545130)': 15,
                'FANCY CHEPNGENO ': 15,
                'IRENE MUNYOKA': 15,
                'EZEKIEL NJOROGE': 15,
            },
            'FGWHHMG0N01': {
                'ANGELA AKIRAPA (Till 4545130)': 40,
                'BEATRICE INGABIRE(Till 4545130)': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE': 20,
            },
            'FGWHHMG0N02': {
                'ANGELA AKIRAPA (Till 4545130)': 40,
                'BEATRICE INGABIRE(Till 4545130)': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE': 20,
            },
            'FGWHTRMG0003': {
                'ANGELA AKIRAPA (Till 4545130)': 40,
                'BEATRICE INGABIRE(Till 4545130)': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE': 20,
            },
        },
    },
    'Jumra': {
        'fsr_sales_targets': {
            'Jedidah Kemunto': 2000000,
            'Ochieng Charles': 1700000,
            'Lenah Cheloti': 1300000,
            'Moses  Ngugi': 3000000,
        },
        'unit_targets': {
            'Jedidah Kemunto': 'Retail',
            'Ochieng Charles': 'Retail',
            'Lenah Cheloti': 'Retail',
            'Moses  Ngugi': 'Marble/Retail',
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
                    'Ochieng Charles': 900000,
                    'Lenah Cheloti': 600000,
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
    'Jumra Eldoret': {
        'fsr_sales_targets': {
            'Jedidah Danyoko': 1824515,
            'Joy Alumasa': 1824515,
            'Miriam Rono': 2606451,
            'Thomas Kiamaiyo': 1824515,
        },
        'unit_targets': {
            'Jedidah Danyoko': 'Retail',
            'Joy Alumasa': 'Retail',
            'Miriam Rono': 'Retail/Marble',
            'Thomas Kiamaiyo': 'Retail',
        },
        'aer_fsr_targets': {
            'Jedidah Danyoko': 500,
            'Joy Alumasa': 500,
            'Miriam Rono': 500,
            'Thomas Kiamaiyo': 500,
        },
        'sub_companies': {
            'Kimberly Clark': {
                'brands': ['Huggies', 'Kotex'],
                'fsr_targets': {
                    'Jedidah Danyoko': 785806,
                    'Joy Alumasa': 785806,
                    'Miriam Rono': 1122581,
                    'Thomas Kiamaiyo': 785806,
                },
                'eco_targets': {
                    'Jedidah Danyoko': 250,
                    'Joy Alumasa': 250,
                    'Miriam Rono': 250,
                    'Thomas Kiamaiyo': 250,
                },
                'brand_targets': {
                    'Huggies': {
                        'Jedidah Danyoko': 100,
                        'Joy Alumasa': 100,
                        'Miriam Rono': 100,
                        'Thomas Kiamaiyo': 100,
                    },
                    'Kotex': {
                        'Jedidah Danyoko': 130,
                        'Joy Alumasa': 130,
                        'Miriam Rono': 130,
                        'Thomas Kiamaiyo': 130,
                    },
                },
            },
            'Dabur': {
                'brands': ['Amla', 'Hobby', 'Dabur', 'Vatika', 'Fem', 'ORS'],
                'fsr_targets': {
                    'Jedidah Danyoko': 1038709,
                    'Joy Alumasa': 1038709,
                    'Miriam Rono': 1483870,
                    'Thomas Kiamaiyo': 1038709,
                },
                'eco_targets': {
                    'Jedidah Danyoko': 250,
                    'Joy Alumasa': 250,
                    'Miriam Rono': 250,
                    'Thomas Kiamaiyo': 250,
                },
                'brand_targets': {
                    'Amla': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                    'Hobby': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                    'Dabur': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                    'Vatika': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                    'Fem': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                    'ORS': {
                        'Jedidah Danyoko': 0,
                        'Joy Alumasa': 0,
                        'Miriam Rono': 0,
                        'Thomas Kiamaiyo': 0,
                    },
                },
            },
        },
    },
    'Reckit': {
        'fsr_sales_targets': {
            'SHA1000': 1500000,
            'SHA1001': 1800000,
            'SHA1002': 1200000,
            'SHA1003': 2000000,
            'SHA1004': 2000000,
            'SHA1005': 2000000,
            'SHA1006': 2000000,
            'SHA1007': 2000000,

        },
        'unit_targets': {
            'SHA1000': 'Whsle',
            'SHA1001': 'Retail',
            'SHA1002': 'Retail',
            'SHA1003': 'Retail',
            'SHA1004': 'Retail',
            'SHA1005': 'Retail',
            'SHA1006': 'Retail',
            'SHA1007': 'Retail',
        },
        'eco_targets': {
            'SHA1000': 75,
            'SHA1001': 200,
            'SHA1002': 200,
            'SHA1003': 200,
            'SHA1004': 200,
            'SHA1005': 200,
            'SHA1006': 200,
            'SHA1007': 200,
        },
        'fsr_representatives': {  # New mapping of FSR codes to representative names
            'SHA1000': 'Fidel Nzwili',
            'SHA1001': 'Elsie Ngugi',
            'SHA1002': 'Enock Joseph',
            'SHA1003': 'Elizabeth Kamau',
            'SHA1004': 'Joshua Kimani',
            'SHA1005': 'Moses Bwire',
            'SHA1006': 'Emmanuel Ngetich',
            'SHA1007': 'Jacqueline Bett',
        },
        'brand_targets': {
            'DETTOL': {
                'SHA1000': 172,
                'SHA1001': 172,
                'SHA1002': 172,
                'SHA1003': 172,
                'SHA1004': 172,
                'SHA1005': 172,
                'SHA1006': 172,
                'SHA1007': 172,
            },
            'HARPIC': {
                'SHA1000': 150,
                'SHA1001': 150,
                'SHA1002': 150,
                'SHA1003': 150,
                'SHA1004': 150,
                'SHA1005': 150,
                'SHA1006': 150,
                'SHA1007': 150,
            },
            'AIRWICK': {
                'SHA1000': 30,
                'SHA1001': 30,
                'SHA1002': 30,
                'SHA1003': 30,
                'SHA1004': 30,
                'SHA1005': 30,
                'SHA1006': 30,
                'SHA1007': 30,

            },
            'DETTOL_SOAP_90G': {  # New entry for Dettol Soap 90g variants
                'SHA1000': 100,    # Example targets; adjust as needed
                'SHA1001': 100,
                'SHA1002': 100,
                'SHA1003': 100,
                'SHA1004': 100,
                'SHA1005': 100,
                'SHA1006': 100,
                'SHA1007': 100,
            },
            'HARPIC_100ML': {  # New entry for Harpic 100ml (SKU 3286535)
                'SHA1000': 58,  # Example targets; adjust as needed
                'SHA1001': 58,
                'SHA1002': 58,
                'SHA1003': 58,
                'SHA1004': 58,
                'SHA1005': 58,
                'SHA1006': 58,
                'SHA1007': 58,
            },
            'JIK': {  # New entry for all Jik SKUs
                'SHA1000': 172,  # Example targets; adjust as needed
                'SHA1001': 172,
                'SHA1002': 172,
                'SHA1003': 172,
                'SHA1004': 172,
                'SHA1005': 172,
                'SHA1006': 172,
                'SHA1007': 172,
            },
            'MORTEIN_DOOM_100ML': {  # New entry for SKU 349130
                'SHA1000': 60,   # Example targets; adjust as needed
                'SHA1001': 60,
                'SHA1002': 60,
                'SHA1003': 60,
                'SHA1004': 60,
                'SHA1005': 60,
                'SHA1006': 60,
                'SHA1007': 60,
            },
            'MORTEIN': {  # New entry for all Mortein Doom products
                'SHA1000': 65,  # Example targets; adjust as needed
                'SHA1001': 65,
                'SHA1002': 65,
                'SHA1003': 65,
                'SHA1004': 65,
                'SHA1005': 65,
                'SHA1006': 65,
                'SHA1007': 65,
            },
            
             
        },
    },
}

def add_totals_row(df):
    try:
        logger.info(f"Adding totals row to DataFrame with columns: {list(df.columns)}")
        logger.info(f"DataFrame shape before adding totals: {df.shape}")

        if df.empty:
            logger.warning("DataFrame is empty; cannot add totals row.")
            return df

        # Identify numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns
        logger.info(f"Numeric columns: {list(numeric_cols)}")

        # Calculate totals for numeric columns
        totals = df[numeric_cols].sum()
        totals_row = pd.DataFrame([totals], columns=numeric_cols)

        # Specifically recalculate % Sales and % ECO for the totals row
        if 'Sales Actual' in totals_row.columns and 'Sales Target' in totals_row.columns:
            sales_actual = totals_row['Sales Actual'].iloc[0]
            sales_target = totals_row['Sales Target'].iloc[0]
            totals_row['% Sales'] = (sales_actual / sales_target * 100) if sales_target != 0 else 0

        if 'ECO Actual' in totals_row.columns and 'ECO Target' in totals_row.columns:
            eco_actual = totals_row['ECO Actual'].iloc[0]
            eco_target = totals_row['ECO Target'].iloc[0]
            if '% ECO' in df.columns:
                totals_row['% ECO'] = (eco_actual / eco_target * 100) if eco_target != 0 else 0
            elif 'ECO %' in df.columns:
                totals_row['ECO %'] = (eco_actual / eco_target * 100) if eco_target != 0 else 0

        # Add non-numeric columns
        non_numeric_cols = df.select_dtypes(exclude=['number']).columns
        logger.info(f"Non-numeric columns: {list(non_numeric_cols)}")
        for col in non_numeric_cols:
            if col in ['FSR', 'Representative']:
                totals_row[col] = 'KD Totals'
            else:
                totals_row[col] = ''

        # Reorder columns to match the original DataFrame
        totals_row = totals_row[df.columns]
        logger.info(f"Totals row created: {totals_row.to_dict()}")

        # Append the totals row to the DataFrame
        df = pd.concat([df, totals_row], ignore_index=True)
        logger.info(f"DataFrame shape after adding totals: {df.shape}")

        return df

    except Exception as e:
        logger.error(f"Error adding totals row: {e}")
        return df
    
def process_canon(file_path, company, df):
    try:
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]

        # General Sales Report
        sales_report = df.groupby('FSR', as_index=False).agg({
            'Amount': 'sum',
            'Customer': pd.Series.nunique
        })
        sales_report.rename(columns={'Amount': 'Sales Actual', 'Customer': 'ECO Actual'}, inplace=True)

        sales_report['Sales Target'] = sales_report['FSR'].map(company_config['fsr_sales_targets']).fillna(0)
        sales_report['Unit'] = sales_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
        sales_report['ECO Target'] = sales_report['FSR'].map(company_config['aer_fsr_targets']).fillna(0)
        sales_report['Sales Balance'] = sales_report['Sales Actual'] - sales_report['Sales Target']
        sales_report['% Sales'] = sales_report.apply(
            lambda row: (row['Sales Actual'] / row['Sales Target']) * 100 if row['Sales Target'] != 0 else 0, axis=1
        )
        sales_report['ECO Balance'] = sales_report['ECO Actual'] - sales_report['ECO Target']
        sales_report['% ECO'] = sales_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )

        sales_report = sales_report[[
            'FSR', 'Unit', 'Sales Target', 'Sales Actual', 'Sales Balance', '% Sales',
            'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
        ]]

        sales_report['% Sales'] = pd.to_numeric(sales_report['% Sales'], errors='coerce')
        sales_report['% ECO'] = pd.to_numeric(sales_report['% ECO'], errors='coerce')
        sales_report = add_totals_row(sales_report)

        # Brand ECO Reports
        brand_reports = {}
        brand_targets = company_config.get('brand_targets', {})
        sku_list = ['FGWHHMG0N01', 'FGWHHMG0N02', 'FGWHTRMG0003']
        for brand in brand_targets.keys():
            if brand not in sku_list:
                brand_df = df[df['Brand'] == brand]
                eco_report = pd.DataFrame({'FSR': list(brand_targets[brand].keys())})
                if not brand_df.empty:
                    actuals = brand_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
                    eco_report = eco_report.merge(actuals, on='FSR', how='left').fillna({'Customer': 0})
                else:
                    eco_report['Customer'] = 0
                eco_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
                eco_report['ECO Target'] = eco_report['FSR'].map(brand_targets.get(brand, {})).fillna(0)
                eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
                eco_report['% ECO'] = eco_report.apply(
                    lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
                )
                eco_report['% ECO'] = pd.to_numeric(eco_report['% ECO'], errors='coerce')
                eco_report = add_totals_row(eco_report)
                brand_reports[brand] = eco_report

        # SKU ECO Reports
        sku_eco_reports = {}
        specific_skus = ['FGWHHMG0N01', 'FGWHHMG0N02', 'FGWHTRMG0003']
        for sku in specific_skus:
            sku_df = df[df['SKU_Code'] == sku]
            eco_report = pd.DataFrame({'FSR': list(brand_targets.get(sku, {}).keys())})
            if not sku_df.empty:
                actuals = sku_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
                eco_report = eco_report.merge(actuals, on='FSR', how='left').fillna({'Customer': 0})
            else:
                eco_report['Customer'] = 0
            eco_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
            eco_report['ECO Target'] = eco_report['FSR'].map(brand_targets.get(sku, {})).fillna(0)
            eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
            eco_report['ECO %'] = eco_report.apply(
                lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
            )
            eco_report['ECO %'] = pd.to_numeric(eco_report['ECO %'], errors='coerce')
            eco_report = eco_report[['FSR', 'ECO Target', 'ECO Actual', 'ECO Balance', 'ECO %']]
            eco_report = add_totals_row(eco_report)
            sku_eco_reports[sku] = eco_report

        return sales_report, brand_reports, sku_eco_reports

    except Exception as e:
        logger.error(f"Error processing {company} file: {e}")
        raise ValueError(f"Error processing the {company} file: {e}")

def process_jumra(file_path, company, df):
    try:
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]

        # General Sales Report
        sales_report = df.groupby('FSR', as_index=False).agg({
            'Amount': 'sum',
            'Customer': pd.Series.nunique
        })
        sales_report.rename(columns={'Amount': 'Sales Actual', 'Customer': 'ECO Actual'}, inplace=True)

        sales_report['Sales Target'] = sales_report['FSR'].map(company_config['fsr_sales_targets']).fillna(0)
        sales_report['Unit'] = sales_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
        sales_report['ECO Target'] = sales_report['FSR'].map(company_config['aer_fsr_targets']).fillna(0)

        sales_report['Sales Balance'] = sales_report['Sales Actual'] - sales_report['Sales Target']
        sales_report['% Sales'] = sales_report.apply(
            lambda row: (row['Sales Actual'] / row['Sales Target']) * 100 if row['Sales Target'] != 0 else 0, axis=1
        )
        sales_report['ECO Balance'] = sales_report['ECO Actual'] - sales_report['ECO Target']
        sales_report['% ECO'] = sales_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )

        sales_report = sales_report[[
            'FSR', 'Unit', 'Sales Target', 'Sales Actual', 'Sales Balance', '% Sales',
            'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
        ]]

        sales_report['% Sales'] = pd.to_numeric(sales_report['% Sales'], errors='coerce')
        sales_report['% ECO'] = pd.to_numeric(sales_report['% ECO'], errors='coerce')
        sales_report = add_totals_row(sales_report)

        # Sub-Company Reports
        sub_company_reports = {}
        for sub_company, config in company_config['sub_companies'].items():
            sub_company_brands = config['brands']
            sub_company_targets = config['fsr_targets']
            sub_company_eco_targets = config.get('eco_targets', {})

            sub_company_df = df[df['Brand'].isin([b.upper() for b in sub_company_brands])]
            if not sub_company_df.empty:
                sub_company_report = sub_company_df.groupby('FSR', as_index=False).agg({
                    'Amount': 'sum',
                    'Customer': pd.Series.nunique
                })
                sub_company_report.rename(columns={'Amount': 'Actual Sales', 'Customer': 'ECO Actual'}, inplace=True)
                sub_company_report['Sales Target'] = sub_company_report['FSR'].map(sub_company_targets).fillna(0)
                sub_company_report['ECO Target'] = sub_company_report['FSR'].map(sub_company_eco_targets).fillna(0)
                sub_company_report['Unit'] = sub_company_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
                sub_company_report['Sales Balance'] = sub_company_report['Actual Sales'] - sub_company_report['Sales Target']
                sub_company_report['% Sales'] = sub_company_report.apply(
                    lambda row: (row['Actual Sales'] / row['Sales Target']) * 100 if row['Sales Target'] != 0 else 0, axis=1
                )
                sub_company_report['ECO Balance'] = sub_company_report['ECO Actual'] - sub_company_report['ECO Target']
                sub_company_report['% ECO'] = sub_company_report.apply(
                    lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
                )
                sub_company_report = sub_company_report[[
                    'FSR', 'Unit', 'Sales Target', 'Actual Sales', 'Sales Balance', '% Sales',
                    'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
                ]]
                sub_company_report['% Sales'] = pd.to_numeric(sub_company_report['% Sales'], errors='coerce')
                sub_company_report['% ECO'] = pd.to_numeric(sub_company_report['% ECO'], errors='coerce')
                sub_company_report = add_totals_row(sub_company_report)
                sub_company_reports[sub_company] = sub_company_report
            else:
                sub_company_report = pd.DataFrame({'FSR': list(sub_company_targets.keys())})
                sub_company_report['Actual Sales'] = 0
                sub_company_report['ECO Actual'] = 0
                sub_company_report['Sales Target'] = sub_company_report['FSR'].map(sub_company_targets).fillna(0)
                sub_company_report['ECO Target'] = sub_company_report['FSR'].map(sub_company_eco_targets).fillna(0)
                sub_company_report['Unit'] = sub_company_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
                sub_company_report['Sales Balance'] = sub_company_report['Actual Sales'] - sub_company_report['Sales Target']
                sub_company_report['% Sales'] = sub_company_report.apply(
                    lambda row: (row['Actual Sales'] / row['Sales Target']) * 100 if row['Sales Target'] != 0 else 0, axis=1
                )
                sub_company_report['ECO Balance'] = sub_company_report['ECO Actual'] - sub_company_report['ECO Target']
                sub_company_report['% ECO'] = sub_company_report.apply(
                    lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
                )
                sub_company_report = sub_company_report[[
                    'FSR', 'Unit', 'Sales Target', 'Actual Sales', 'Sales Balance', '% Sales',
                    'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
                ]]
                sub_company_report['% Sales'] = pd.to_numeric(sub_company_report['% Sales'], errors='coerce')
                sub_company_report['% ECO'] = pd.to_numeric(sub_company_report['% ECO'], errors='coerce')
                sub_company_report = add_totals_row(sub_company_report)
                sub_company_reports[sub_company] = sub_company_report

        # ECO Reports for Brands
        eco_reports = {}
        for sub_company, config in company_config['sub_companies'].items():
            brand_targets = config.get('brand_targets', {})
            for brand, targets in brand_targets.items():
                brand_df = df[df['Brand'] == brand.upper()]
                eco_report = pd.DataFrame({'FSR': list(targets.keys())})
                if not brand_df.empty:
                    actuals = brand_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
                    eco_report = eco_report.merge(actuals, on='FSR', how='left').fillna({'Customer': 0})
                else:
                    eco_report['Customer'] = 0
                eco_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
                eco_report['ECO Target'] = eco_report['FSR'].map(targets).fillna(0)
                eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
                eco_report['% ECO'] = eco_report.apply(
                    lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
                )
                eco_report['% ECO'] = pd.to_numeric(eco_report['% ECO'], errors='coerce')
                eco_report = add_totals_row(eco_report)
                eco_reports[brand] = eco_report

        return sales_report, sub_company_reports, eco_reports

    except Exception as e:
        logger.error(f"Error processing {company} file: {e}")
        raise ValueError(f"Error processing the {company} file: {e}")

def process_reckit(file_path, company, df):
    try:
        # Rename columns to match internal format
        df = df.rename(columns={
            'Route Code': 'FSR',
            'Total Net Amount': 'Amount',
            'Customer Code': 'Customer',
            'Product Description': 'Brand',
            'Product Code': 'SKU_Code'
        })
        
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]
        fsr_reps = company_config.get('fsr_representatives', {})

        # General Sales and ECO Report per Route Code
        sales_report = df.groupby('FSR', as_index=False).agg({
            'Amount': 'sum',
            'Customer': pd.Series.nunique
        })
        sales_report.rename(columns={'Amount': 'Sales Actual', 'Customer': 'ECO Actual'}, inplace=True)

        sales_report['Sales Target'] = sales_report['FSR'].map(company_config['fsr_sales_targets']).fillna(0)
        sales_report['Unit'] = sales_report['FSR'].map(company_config['unit_targets']).fillna('Unknown')
        sales_report['ECO Target'] = sales_report['FSR'].map(company_config['eco_targets']).fillna(0)
        sales_report['Sales Balance'] = sales_report['Sales Actual'] - sales_report['Sales Target']
        sales_report['% Sales'] = sales_report.apply(
            lambda row: (row['Sales Actual'] / row['Sales Target']) * 100 if row['Sales Target'] != 0 else 0, axis=1
        )
        sales_report['ECO Balance'] = sales_report['ECO Actual'] - sales_report['ECO Target']
        sales_report['% ECO'] = sales_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )

        # Replace FSR with representative name and rename the column
        sales_report['FSR'] = sales_report['FSR'].apply(
            lambda x: fsr_reps.get(x, 'Unknown') if x in fsr_reps else x
        )
        sales_report.rename(columns={'FSR': 'Representative'}, inplace=True)

        sales_report = sales_report[[
            'Representative', 'Unit', 'Sales Target', 'Sales Actual', 'Sales Balance', '% Sales',
            'ECO Target', 'ECO Actual', 'ECO Balance', '% ECO'
        ]]
        sales_report['% Sales'] = pd.to_numeric(sales_report['% Sales'], errors='coerce')
        sales_report['% ECO'] = pd.to_numeric(sales_report['% ECO'], errors='coerce')
        sales_report = add_totals_row(sales_report)

        # Brand ECO Reports per Route Code
        brand_reports = {}
        brand_targets = company_config.get('brand_targets', {})
        standard_brands = ['DETTOL', 'HARPIC', 'AIRWICK', 'JIK', 'MORTEIN']
        for brand in standard_brands:
            if brand == 'DETTOL':
                brand_df = df[df['Brand'].str.contains('DETTOL SOAP', case=False, na=False)]
            elif brand == 'MORTEIN':
                brand_df = df[df['Brand'].str.contains('MORTEIN DOOM', case=False, na=False)]
            else:
                brand_df = df[df['Brand'].str.contains(brand, case=False, na=False)]
            eco_report = pd.DataFrame({'FSR': list(brand_targets[brand].keys())})
            if not brand_df.empty:
                actuals = brand_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
                eco_report = eco_report.merge(actuals, on='FSR', how='left').fillna({'Customer': 0})
            else:
                eco_report['Customer'] = 0
            eco_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
            eco_report['ECO Target'] = eco_report['FSR'].map(brand_targets.get(brand, {})).fillna(0)
            eco_report['ECO Balance'] = eco_report['ECO Actual'] - eco_report['ECO Target']
            eco_report['% ECO'] = eco_report.apply(
                lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
            )
            # Replace FSR with representative name and rename the column
            eco_report['FSR'] = eco_report['FSR'].apply(
                lambda x: fsr_reps.get(x, 'Unknown') if x in fsr_reps else x
            )
            eco_report.rename(columns={'FSR': 'Representative'}, inplace=True)
            eco_report['% ECO'] = pd.to_numeric(eco_report['% ECO'], errors='coerce')
            eco_report = add_totals_row(eco_report)
            brand_reports[brand] = eco_report

        # Specific ECO Tracking for Dettol Soap 90g Variants
        dettol_soap_90g_codes = [
            '3325468', '366184', '8149556', '3244641', '366195', 
            '3244643', '366159', '8149555', '3244645', '366177', '3244644','3298961'
        ]
        dettol_soap_df = df[df['SKU_Code'].isin(dettol_soap_90g_codes)]
        dettol_soap_90g_report = pd.DataFrame({'FSR': list(brand_targets['DETTOL_SOAP_90G'].keys())})
        if not dettol_soap_df.empty:
            dettol_actuals = dettol_soap_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
            dettol_soap_90g_report = dettol_soap_90g_report.merge(dettol_actuals, on='FSR', how='left').fillna({'Customer': 0})
        else:
            dettol_soap_90g_report['Customer'] = 0
        dettol_soap_90g_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
        dettol_soap_90g_report['ECO Target'] = dettol_soap_90g_report['FSR'].map(brand_targets['DETTOL_SOAP_90G']).fillna(0)
        dettol_soap_90g_report['ECO Balance'] = dettol_soap_90g_report['ECO Actual'] - dettol_soap_90g_report['ECO Target']
        dettol_soap_90g_report['% ECO'] = dettol_soap_90g_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )
        # Replace FSR with representative name and rename the column
        dettol_soap_90g_report['FSR'] = dettol_soap_90g_report['FSR'].apply(
            lambda x: fsr_reps.get(x, 'Unknown') if x in fsr_reps else x
        )
        dettol_soap_90g_report.rename(columns={'FSR': 'Representative'}, inplace=True)
        dettol_soap_90g_report['% ECO'] = pd.to_numeric(dettol_soap_90g_report['% ECO'], errors='coerce')
        dettol_soap_90g_report = add_totals_row(dettol_soap_90g_report)
        brand_reports['DETTOL_SOAP_90G'] = dettol_soap_90g_report

        # Specific ECO Tracking for Mortein Doom 100ml (SKU 349130)
        mortein_doom_100ml_code = ['349130']
        mortein_doom_df = df[df['SKU_Code'].isin(mortein_doom_100ml_code)]
        mortein_doom_100ml_report = pd.DataFrame({'FSR': list(brand_targets['MORTEIN_DOOM_100ML'].keys())})
        if not mortein_doom_df.empty:
            mortein_actuals = mortein_doom_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
            mortein_doom_100ml_report = mortein_doom_100ml_report.merge(mortein_actuals, on='FSR', how='left').fillna({'Customer': 0})
        else:
            mortein_doom_100ml_report['Customer'] = 0
        mortein_doom_100ml_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
        mortein_doom_100ml_report['ECO Target'] = mortein_doom_100ml_report['FSR'].map(brand_targets['MORTEIN_DOOM_100ML']).fillna(0)
        mortein_doom_100ml_report['ECO Balance'] = mortein_doom_100ml_report['ECO Actual'] - mortein_doom_100ml_report['ECO Target']
        mortein_doom_100ml_report['% ECO'] = mortein_doom_100ml_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )
        # Replace FSR with representative name and rename the column
        mortein_doom_100ml_report['FSR'] = mortein_doom_100ml_report['FSR'].apply(
            lambda x: fsr_reps.get(x, 'Unknown') if x in fsr_reps else x
        )
        mortein_doom_100ml_report.rename(columns={'FSR': 'Representative'}, inplace=True)
        mortein_doom_100ml_report['% ECO'] = pd.to_numeric(mortein_doom_100ml_report['% ECO'], errors='coerce')
        mortein_doom_100ml_report = add_totals_row(mortein_doom_100ml_report)
        brand_reports['MORTEIN_DOOM_100ML'] = mortein_doom_100ml_report

        # Specific ECO Tracking for Harpic 100ml (Placeholder SKU 123456)
        harpic_100ml_code = ['123456']  # Replace with actual SKU for Harpic 100ml
        logger.info(f"Filtering Harpic 100ml with SKU: {harpic_100ml_code}")
        harpic_100ml_df = df[df['SKU_Code'].isin(harpic_100ml_code)]
        logger.info(f"Harpic 100ml DataFrame size: {len(harpic_100ml_df)} rows")
        if not harpic_100ml_df.empty:
            logger.info(f"Harpic 100ml data found: {harpic_100ml_df[['SKU_Code', 'Brand', 'FSR']].head().to_dict()}")
        harpic_100ml_report = pd.DataFrame({'FSR': list(brand_targets['HARPIC_100ML'].keys())})
        if not harpic_100ml_df.empty:
            harpic_actuals = harpic_100ml_df.groupby('FSR', as_index=False).agg({'Customer': pd.Series.nunique})
            harpic_100ml_report = harpic_100ml_report.merge(harpic_actuals, on='FSR', how='left').fillna({'Customer': 0})
        else:
            harpic_100ml_report['Customer'] = 0
        harpic_100ml_report.rename(columns={'Customer': 'ECO Actual'}, inplace=True)
        harpic_100ml_report['ECO Target'] = harpic_100ml_report['FSR'].map(brand_targets['HARPIC_100ML']).fillna(0)
        harpic_100ml_report['ECO Balance'] = harpic_100ml_report['ECO Actual'] - harpic_100ml_report['ECO Target']
        harpic_100ml_report['% ECO'] = harpic_100ml_report.apply(
            lambda row: (row['ECO Actual'] / row['ECO Target']) * 100 if row['ECO Target'] != 0 else 0, axis=1
        )
        # Replace FSR with representative name and rename the column
        harpic_100ml_report['FSR'] = harpic_100ml_report['FSR'].apply(
            lambda x: fsr_reps.get(x, 'Unknown') if x in fsr_reps else x
        )
        harpic_100ml_report.rename(columns={'FSR': 'Representative'}, inplace=True)
        harpic_100ml_report['% ECO'] = pd.to_numeric(harpic_100ml_report['% ECO'], errors='coerce')
        harpic_100ml_report = add_totals_row(harpic_100ml_report)
        logger.info(f"Harpic 100ml Report generated: {harpic_100ml_report.to_dict()}")
        brand_reports['HARPIC_100ML'] = harpic_100ml_report

        return sales_report, brand_reports

    except Exception as e:
        logger.error(f"Error processing Reckit file: {e}")
        raise ValueError(f"Error processing the Reckit file: {e}")
    
def process_file(file_path, company):
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension == '.xlsx':
        date_range_df = pd.read_excel(file_path, nrows=1, header=None)
        date_range = date_range_df.iloc[0, 0]
        df = pd.read_excel(file_path, skiprows=1)
    elif file_extension == '.csv':
        df = pd.read_csv(file_path)
        current_date = datetime.today()
        month_start = current_date.replace(day=1).strftime('%Y-%m-%d')
        month_end = (current_date.replace(day=1, month=current_date.month % 12 + 1) - pd.Timedelta(days=1)).strftime('%Y-%m-%d')
        date_range = f"From: {month_start} To: {month_end}"
        df.columns = df.columns.str.strip()
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

    if company in ['Canon', 'Canon Eldoret']:
        sales_report, brand_reports, sku_eco_reports = process_canon(file_path, company, df)
        return sales_report, brand_reports, sku_eco_reports, date_range
    elif company in ['Jumra', 'Jumra Eldoret']:
        sales_report, sub_company_reports, eco_reports = process_jumra(file_path, company, df)
        return sales_report, sub_company_reports, eco_reports, date_range
    elif company == 'Reckit':
        sales_report, brand_reports = process_reckit(file_path, company, df)
        return sales_report, brand_reports, None, date_range
    else:
        raise ValueError(f"Invalid company: {company}")

def create_consolidated_excel(company, file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, skiprows=1)
    elif file_extension == '.csv':
        df = pd.read_csv(file_path)
        if company == 'Reckit':
            df = df.rename(columns={
                'Route Code': 'FSR',
                'Total Net Amount': 'Amount',
                'Customer Code': 'Customer',
                'Product Description': 'Brand',
                'Product Code': 'SKU_Code'
            })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        start_row = 0

        if company in ['Canon', 'Canon Eldoret']:
            sales_report, brand_reports, sku_eco_reports = process_canon(file_path, company, df)
            # Adjust percentage columns for Excel
            if '% Sales' in sales_report.columns:
                sales_report['% Sales'] = sales_report['% Sales'] / 100
            if '% ECO' in sales_report.columns:
                sales_report['% ECO'] = sales_report['% ECO'] / 100
            sales_report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
            start_row += len(sales_report) + 2
            for brand, report in brand_reports.items():
                if report is not None and not report.empty:
                    if '% ECO' in report.columns:
                        report['% ECO'] = report['% ECO'] / 100
                    report.insert(0, 'Report Type', f"{brand} ECO")
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2
            for sku, report in sku_eco_reports.items():
                if report is not None and not report.empty:
                    if '% ECO' in report.columns:
                        report['% ECO'] = report['% ECO'] / 100
                    report.insert(0, 'Report Type', f"{sku} ECO")
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2
        elif company in ['Jumra', 'Jumra Eldoret']:
            sales_report, sub_company_reports, eco_reports = process_jumra(file_path, company, df)
            # Adjust percentage columns for Excel
            if '% Sales' in sales_report.columns:
                sales_report['% Sales'] = sales_report['% Sales'] / 100
            if '% ECO' in sales_report.columns:
                sales_report['% ECO'] = sales_report['% ECO'] / 100
            sales_report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
            start_row += len(sales_report) + 2
            for sub_company, report in sub_company_reports.items():
                if report is not None and not report.empty:
                    if '% Sales' in report.columns:
                        report['% Sales'] = report['% Sales'] / 100
                    if '% ECO' in report.columns:
                        report['% ECO'] = report['% ECO'] / 100
                    report.insert(0, 'Report Type', f"{sub_company} Sales")
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2
            for brand, report in eco_reports.items():
                if report is not None and not report.empty:
                    if '% ECO' in report.columns:
                        report['% ECO'] = report['% ECO'] / 100
                    report.insert(0, 'Report Type', f"{brand} ECO")
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2
        elif company == 'Reckit':
            sales_report, brand_reports = process_reckit(file_path, company, df)
            # Adjust percentage columns for Excel
            if '% Sales' in sales_report.columns:
                sales_report['% Sales'] = sales_report['% Sales'] / 100
            if '% ECO' in sales_report.columns:
                sales_report['% ECO'] = sales_report['% ECO'] / 100
            sales_report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
            start_row += len(sales_report) + 2
            for brand, report in brand_reports.items():
                if report is not None and not report.empty:
                    if '% ECO' in report.columns:
                        report['% ECO'] = report['% ECO'] / 100
                    if '% Sales' in report.columns:
                        report['% Sales'] = report['% Sales'] / 100
                    if brand == 'DETTOL_SOAP_90G':
                        report_type = "DETTOL SOAP 90G ECO"
                    elif brand == 'MORTEIN_DOOM_100ML':
                        report_type = "MORTEIN DOOM 100ML ECO"
                    elif brand == 'HARPIC_100ML':
                        report_type = "HARPIC 100ML ECO"
                    elif brand == 'DETTOL':
                        report_type = "DETTOL SOAP ECO"
                    elif brand == 'MORTEIN':
                        report_type = "MORTEIN DOOM ECO"
                    else:
                        report_type = f"{brand} ECO"
                    report.insert(0, 'Report Type', report_type)
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2

        # Apply formatting to the Excel sheet
        sheet = writer.sheets['Consolidated Report']
        for col in sheet.columns:
            col_letter = col[0].column_letter
            header = col[0].value  # Get the column header
            # Apply number formatting based on column header
            for cell in col[1:]:  # Skip the header row
                if header in ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    cell.number_format = '#,##0.00'
                elif header in ['% Sales', '% ECO']:
                    cell.number_format = '0.00%'

    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def upload_file():

    logger.debug("Entering upload_file route")
    if request.method == 'POST':
        logger.debug("POST request received")
        if 'file' not in request.files:
            logger.error("No file part in request")
            return 'No file uploaded.', 400

        file = request.files['file']
        company = request.form.get('company')
        logger.debug(f"Company selected: {company}, File: {file.filename}")

        if file.filename == '':
            logger.error("No file selected")
            return 'No file selected.', 400

        if not company:
            logger.error("No company selected")
            return 'No company selected.', 400

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        session['uploaded_file_path'] = file_path
        logger.debug(f"File saved at: {file_path}")

        try:
            logger.debug("Processing file")
            reports_data = process_file(file_path, company)
            logger.debug(f"Reports data processed: {reports_data}")
            today_date = datetime.today().strftime('%Y-%m-%d')

            def format_percentage(val):
                if not isinstance(val, (int, float)) or pd.isna(val):
                    return val
                formatted_val = f"{val:.2f}%"
                if val < 50:
                    return f'<span class="percent-red">{formatted_val}</span>'
                elif 50 <= val < 90:
                    return f'<span class="percent-light-red">{formatted_val}</span>'
                elif 90 <= val < 100:
                    return f'<span class="percent-light-green">{formatted_val}</span>'
                else:
                    return f'<span class="percent-green">{formatted_val}</span>'

            def format_number(val):
                if isinstance(val, (int, float)) and not pd.isna(val):
                    return f"{val:,.2f}"
                return val

            if company in ['Canon', 'Canon Eldoret']:
                logger.debug("Processing Canon report")
                sales_report, brand_reports, sku_eco_reports, date_range = reports_data
                sales_report_html = sales_report.copy()
                for col in ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    sales_report_html[col] = sales_report_html[col].apply(format_number)
                for col in ['% Sales', '% ECO']:
                    sales_report_html[col] = sales_report_html[col].apply(format_percentage)
                html_table = sales_report_html.to_html(
                    classes="table table-striped table-bordered", index=False, escape=False
                )

                brand_reports_html = {}
                for brand, report in brand_reports.items():
                    report_html = report.copy()
                    for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                        report_html[col] = report_html[col].apply(format_number)
                    report_html['% ECO'] = report_html['% ECO'].apply(format_percentage)
                    brand_reports_html[brand] = report_html.to_html(
                        classes="table table-striped table-bordered", index=False, escape=False
                    )

                sku_reports_html = {}
                for sku, report in sku_eco_reports.items():
                    report_html = report.copy()
                    for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                        report_html[col] = report_html[col].apply(format_number)
                    report_html['ECO %'] = report_html['ECO %'].apply(format_percentage)
                    sku_reports_html[sku] = report_html.to_html(
                        classes="table table-striped table-bordered", index=False, escape=False
                    )

                logger.debug("Rendering Canon report")
                return render_template(
                    "canon_report.html",
                    sales_report=html_table,
                    brand_reports=brand_reports_html,
                    sku_eco_reports=sku_reports_html,
                    today_date=today_date,
                    date_range=date_range,
                    company=company,
                    download_filename=f"{company}_consolidated_report.xlsx",
                    original_filename=file.filename
                )
            
            elif company in ['Jumra', 'Jumra Eldoret']:
                logger.debug("Processing Jumra report")
                sales_report, sub_company_reports, eco_reports, date_range = reports_data
                sales_report_html = sales_report.copy()
                for col in ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    sales_report_html[col] = sales_report_html[col].apply(format_number)
                for col in ['% Sales', '% ECO']:
                    sales_report_html[col] = sales_report_html[col].apply(format_percentage)
                html_table = sales_report_html.to_html(
                    classes="table table-striped table-bordered", index=False, escape=False
                )

                sub_company_reports_html = {}
                for sub_company, report in sub_company_reports.items():
                    report_html = report.copy()
                    for col in ['Sales Target', 'Actual Sales', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                        report_html[col] = report_html[col].apply(format_number)
                    for col in ['% Sales', '% ECO']:
                        report_html[col] = report_html[col].apply(format_percentage)
                    sub_company_reports_html[sub_company] = report_html.to_html(
                        classes="table table-striped table-bordered", index=False, escape=False
                    )

                eco_reports_html = {}
                for brand, report in eco_reports.items():
                    report_html = report.copy()
                    for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                        report_html[col] = report_html[col].apply(format_number)
                    report_html['% ECO'] = report_html['% ECO'].apply(format_percentage)
                    eco_reports_html[brand] = report_html.to_html(
                        classes="table table-striped table-bordered", index=False, escape=False
                    )

                logger.debug("Rendering Jumra report")
                return render_template(
                    "jumra_report.html",
                    sales_report=html_table,
                    sub_company_reports=sub_company_reports_html,
                    eco_reports=eco_reports_html,
                    today_date=today_date,
                    date_range=date_range,
                    company=company,
                    company_config=COMPANIES[company]['sub_companies'],
                    download_filename=f"{company}_consolidated_report.xlsx",
                    original_filename=file.filename
                )
            
            elif company == 'Reckit':
                logger.debug("Processing Reckit report")
                sales_report, brand_reports, _, date_range = reports_data
                logger.debug(f"Reckit sales_report: {sales_report.to_dict()}")
                logger.debug(f"Reckit brand_reports: { {k: v.to_dict() for k, v in brand_reports.items()} }")
                sales_report_html = sales_report.copy()
                for col in ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    sales_report_html[col] = sales_report_html[col].apply(format_number)
                for col in ['% Sales', '% ECO']:
                    sales_report_html[col] = sales_report_html[col].apply(format_percentage)
                html_table = sales_report_html.to_html(
                    classes="table table-striped table-bordered", index=False, escape=False
                )

                brand_reports_html = {}
                for brand, report in brand_reports.items():
                    report_html = report.copy()
                    for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                        report_html[col] = report_html[col].apply(format_number)
                    report_html['% ECO'] = report_html['% ECO'].apply(format_percentage)
                    brand_reports_html[brand] = report_html.to_html(
                        classes="table table-striped table-bordered", index=False, escape=False
                    )

                logger.debug("Rendering Reckit report")
                return render_template(
                    "reckit_report.html",
                    sales_report=html_table,
                    brand_reports=brand_reports_html,
                    today_date=today_date,
                    date_range=date_range,
                    company=company,
                    download_filename=f"{company}_consolidated_report.csv",
                    original_filename=file.filename
                )

        except Exception as e:
            logger.error(f"Upload error: {str(e)}", exc_info=True)
            return f'Error: {str(e)}', 500

    logger.debug("Rendering upload.html")
    return render_template('upload.html')

@app.route('/download_consolidated/<company>/<filename>')
def download_consolidated(company, filename):
    file_path = session.get('uploaded_file_path')
    if file_path and os.path.exists(file_path):
        try:
            excel_file = create_consolidated_excel(company, file_path)
            return send_file(
                excel_file,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=filename
            )
        except Exception as e:
            logger.error(f"Error generating consolidated report: {e}")
            return f"Error generating consolidated report: {e}", 500
    return "No uploaded file available. Please upload a file first.", 404

@app.route('/download_original/<company>')
def download_original(company):
    file_path = session.get('uploaded_file_path')
    if file_path and os.path.exists(file_path):
        original_filename = os.path.basename(file_path)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if file_path.endswith('.xlsx') else 'text/csv'
        response = send_file(
            file_path,
            mimetype=mime_type,
            as_attachment=True,
            download_name=original_filename
        )
        try:
            os.remove(file_path)
            session.pop('uploaded_file_path', None)
        except:
            pass
        return response
    return "File not found or already downloaded.", 404

if __name__ == '__main__':
    app.run(debug=True)