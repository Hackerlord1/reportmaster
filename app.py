from flask import Flask, render_template, request, send_file, session
from openpyxl.styles import Font, PatternFill, numbers
from openpyxl.styles import numbers
import pandas as pd
import os
from datetime import datetime
import io
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Replace with a secure key

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
            'ANGELA AKIRAPA': 1740000,
            'BEATRICE INGABIRE': 2784000,
            'FANCY CHEPNGENO ': 928000,
            'IRENE MUNYOKA': 812000,
            'EZEKIEL NJOROGE': 1044000,
            'Walkin  Sales': 3132000.
        },
        'unit_targets': {
            'ANGELA AKIRAPA': 'Retail',
            'BEATRICE INGABIRE': 'Retail/Marble',
            'FANCY CHEPNGENO ': 'Retail',
            'IRENE MUNYOKA': 'Retail',
            'EZEKIEL NJOROGE': 'Ratail',
            'Walkin  Sales': 'Wholesale',
        },
        'aer_fsr_targets': {
            'ANGELA AKIRAPA': 150,
            'BEATRICE INGABIRE': 150,
            'FANCY CHEPNGENO ': 250,
            'IRENE MUNYOKA': 250,
            'EZEKIEL NJOROGE': 250,
            'Walkin  Sales': 23,
        },
        'brand_targets': {
            'GOODNIGHT INSECT KILLER': {
                'ANGELA AKIRAPA': 12,
                'BEATRICE INGABIRE': 12,
                'FANCY CHEPNGENO ': 12,
                'IRENE MUNYOKA': 12,
                'EZEKIEL NJOROGE': 15,
            },
            'AER POWER POCKET': {
                'ANGELA AKIRAPA': 15,
                'BEATRICE INGABIRE': 15,
                'FANCY CHEPNGENO ': 15,
                'IRENE MUNYOKA': 15,
                'EZEKIEL NJOROGE': 15,
            },
            'FGWHHMG0N01': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE': 20,
            },
            'FGWHHMG0N02': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE': 20,
            },
            'FGWHTRMG0003': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
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
}

def add_totals_row(df):
    totals = {}
    for column in df.columns:
        if pd.api.types.is_numeric_dtype(df[column]):
            if column.startswith('%'):
                totals[column] = df[column].mean()
            else:
                totals[column] = df[column].sum()
        elif column == "FSR":
            totals[column] = "KD Totals"
        else:
            totals[column] = ""
    totals_row = pd.DataFrame(totals, index=["KD Totals"])
    return pd.concat([df, totals_row], ignore_index=True)

def process_canon(file_path, company, df):
    """
    Process Canon sales data from an Excel file and generate sales, brand, and SKU reports.
    
    Args:
        file_path (str): Path to the uploaded Excel file (kept for compatibility)
        company (str): Company name ('Canon' or 'Canon Eldoret')
        df (pd.DataFrame): Preloaded DataFrame with data starting from the second row
    
    Returns:
        tuple: (sales_report, brand_reports, sku_eco_reports)
    """
    try:
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]

        # 1. General Sales Report
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

        # 2. Brand ECO Reports
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

        # 3. SKU ECO Reports
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
            eco_report = add_totals_row(eco_report)
            if eco_report['ECO %'].iloc[:-1].sum() == eco_report['ECO %'].iloc[-1]:
                eco_report.loc[eco_report.index[-1], 'ECO %'] = eco_report['ECO %'].iloc[:-1].mean()
            sku_eco_reports[sku] = eco_report

        return sales_report, brand_reports, sku_eco_reports

    except Exception as e:
        logger.error(f"Error processing {company} file: {e}")
        raise ValueError(f"Error processing the {company} file: {e}")

def process_jumra(file_path, company, df):
    """
    Process Jumra sales data from an Excel file and generate sales, sub-company, and ECO reports.
    
    Args:
        file_path (str): Path to the uploaded Excel file (kept for compatibility)
        company (str): Company name ('Jumra' or 'Jumra Eldoret')
        df (pd.DataFrame): Preloaded DataFrame with data starting from the second row
    
    Returns:
        tuple: (sales_report, sub_company_reports, eco_reports)
    """
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

def process_excel(file_path, company):
    # Read the first row to extract the date range
    date_range_df = pd.read_excel(file_path, nrows=1, header=None)
    date_range = date_range_df.iloc[0, 0]  # Assuming "From: ... To: ..." is in the first column

    # Read the actual data, skipping the first row and using the second row as header
    df = pd.read_excel(file_path, skiprows=1)  # Skip the first row, second row becomes header

    if company in ['Canon', 'Canon Eldoret']:
        sales_report, brand_reports, sku_eco_reports = process_canon(file_path, company, df)
        return sales_report, brand_reports, sku_eco_reports, date_range
    elif company in ['Jumra', 'Jumra Eldoret']:
        sales_report, sub_company_reports, eco_reports = process_jumra(file_path, company, df)
        return sales_report, sub_company_reports, eco_reports, date_range
    else:
        raise ValueError(f"Invalid company: {company}")

def create_consolidated_excel(company, file_path):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Read the Excel file, skipping the first row (date range), and use the second row as header
        df = pd.read_excel(file_path, skiprows=1)
        
        # Create a single DataFrame to hold all consolidated data
        consolidated_df = pd.DataFrame()
        start_row = 0

        if company in ['Canon', 'Canon Eldoret']:
            sales_report, brand_reports, sku_eco_reports = process_canon(file_path, company, df)
            
            # Add General Sales Report
            sales_report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
            start_row += len(sales_report) + 2  # Add 2 rows spacing
            
            # Add Brand Reports
            for brand, report in brand_reports.items():
                if report is not None and not report.empty:
                    report.insert(0, 'Report Type', f"{brand} ECO")  # Add identifier column
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2  # Add 2 rows spacing
            
            # Add SKU Reports
            for sku, report in sku_eco_reports.items():
                if report is not None and not report.empty:
                    report.insert(0, 'Report Type', f"{sku} ECO")  # Add identifier column
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2  # Add 2 rows spacing

        elif company in ['Jumra', 'Jumra Eldoret']:
            sales_report, sub_company_reports, eco_reports = process_jumra(file_path, company, df)
            
            # Add General Sales Report
            sales_report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
            start_row += len(sales_report) + 2  # Add 2 rows spacing
            
            # Add Sub-Company Reports
            for sub_company, report in sub_company_reports.items():
                if report is not None and not report.empty:
                    report.insert(0, 'Report Type', f"{sub_company} Sales")  # Add identifier column
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2  # Add 2 rows spacing
            
            # Add ECO Reports
            for brand, report in eco_reports.items():
                if report is not None and not report.empty:
                    report.insert(0, 'Report Type', f"{brand} ECO")  # Add identifier column
                    report.to_excel(writer, sheet_name='Consolidated Report', startrow=start_row, index=False)
                    start_row += len(report) + 2  # Add 2 rows spacing

        # Apply formatting to the entire sheet
        sheet = writer.sheets['Consolidated Report']
        for col in sheet.columns:
            col_letter = col[0].column_letter
            col_idx = col[0].column - 1
            # Since we're now using one sheet, we need to get the column name from the DataFrame headers
            # We'll apply formatting based on common column names across all reports
            for cell in col:
                cell_value = cell.value
                if isinstance(cell_value, str) and any(x in cell_value for x in ['Target', 'Actual', 'Balance']):
                    cell.number_format = '#,##0.00'
                elif isinstance(cell_value, str) and any(x in cell_value for x in ['%', 'ECO %']):
                    cell.number_format = '0.00%'

    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded.', 400

        file = request.files['file']
        company = request.form.get('company')

        if file.filename == '':
            return 'No file selected.', 400

        if not company:
            return 'No company selected.', 400

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        session['uploaded_file_path'] = file_path

        try:
            reports_data = process_excel(file_path, company)
            if company in ['Canon', 'Canon Eldoret']:
                sales_report, brand_reports, sku_eco_reports, date_range = reports_data
            elif company in ['Jumra', 'Jumra Eldoret']:
                sales_report, sub_company_reports, eco_reports, date_range = reports_data

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
        except Exception as e:
            logger.error(f"Upload error: {e}")
            return f'Error: {e}', 500

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
        response = send_file(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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