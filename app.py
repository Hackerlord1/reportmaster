from flask import Flask, render_template, request, send_file, session
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
            'FGWHTRMG0003': {
                'Jane Gichohi': 30,
                'Glayds  Bundotich': 30,
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
            'EZEKIEL NJOROGE':1044000,
            'Walkin  Sales':3132000.
            
        },
        'unit_targets': {
            'ANGELA AKIRAPA': 'Retail',
            'BEATRICE INGABIRE': 'Retail/Marble',
            'FANCY CHEPNGENO ': 'Retail',
            'IRENE MUNYOKA': 'Retail',
            'EZEKIEL NJOROGE':'Ratail',
            'Walkin  Sales': 'Wholesale',
        },
        'aer_fsr_targets': {
            'ANGELA AKIRAPA': 150,
            'BEATRICE INGABIRE': 150,
            'FANCY CHEPNGENO ': 250,
            'IRENE MUNYOKA': 250,
            'EZEKIEL NJOROGE':250,
            
        },
        'brand_targets': {
            'GOODNIGHT INSECT KILLER': {
                'ANGELA AKIRAPA': 12,
                'BEATRICE INGABIRE': 12,
                'FANCY CHEPNGENO ': 12,
                'IRENE MUNYOKA': 12,
                'EZEKIEL NJOROGE':15,
                
            },
            'AER POWER POCKET': {
                'ANGELA AKIRAPA': 15,
                'BEATRICE INGABIRE': 15,
                'FANCY CHEPNGENO ': 15,
                'IRENE MUNYOKA': 15,
                'EZEKIEL NJOROGE':15,
            
            },
            'FGWHHMG0N01': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE':20,
                
            },
            'FGWHHMG0N02': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE':20,
                
            },
            'FGWHTRMG0003': {
                'ANGELA AKIRAPA': 40,
                'BEATRICE INGABIRE': 50,
                'FANCY CHEPNGENO ': 25,
                'IRENE MUNYOKA': 25,
                'EZEKIEL NJOROGE':20,
                
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

def process_canon(file_path, company):
    try:
        df = pd.read_excel(file_path)
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]

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

        sales_report = add_totals_row(sales_report)

        numeric_columns = ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']
        for col in numeric_columns:
            sales_report[col] = sales_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
        percentage_columns = ['% Sales', '% ECO']
        for col in percentage_columns:
            sales_report[col] = sales_report[col].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)

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
                eco_report = add_totals_row(eco_report)
                for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                    eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                eco_report['% ECO'] = eco_report['% ECO'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
                brand_reports[brand] = eco_report

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
            eco_report = add_totals_row(eco_report)
            for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
            eco_report['ECO %'] = eco_report['ECO %'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
            sku_eco_reports[sku] = eco_report

        return sales_report, brand_reports, sku_eco_reports

    except Exception as e:
        logger.error(f"Error processing {company} file: {e}")
        raise ValueError(f"Error processing the {company} file: {e}")

def process_jumra(file_path, company):
    try:
        df = pd.read_excel(file_path)
        df['SKU_Code'] = df['SKU_Code'].astype(str).str.strip().str.upper()
        df['Brand'] = df['Brand'].astype(str).str.strip().str.upper()

        company_config = COMPANIES[company]

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

        sales_report = add_totals_row(sales_report)

        numeric_columns = ['Sales Target', 'Sales Actual', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']
        for col in numeric_columns:
            sales_report[col] = sales_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
        percentage_columns = ['% Sales', '% ECO']
        for col in percentage_columns:
            sales_report[col] = sales_report[col].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)

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
                sub_company_report = add_totals_row(sub_company_report)
                for col in ['Sales Target', 'Actual Sales', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                for col in ['% Sales', '% ECO']:
                    sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
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
                sub_company_report = add_totals_row(sub_company_report)
                for col in ['Sales Target', 'Actual Sales', 'Sales Balance', 'ECO Target', 'ECO Actual', 'ECO Balance']:
                    sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                for col in ['% Sales', '% ECO']:
                    sub_company_report[col] = sub_company_report[col].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
                sub_company_reports[sub_company] = sub_company_report

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
                eco_report = add_totals_row(eco_report)
                for col in ['ECO Target', 'ECO Actual', 'ECO Balance']:
                    eco_report[col] = eco_report[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                eco_report['% ECO'] = eco_report['% ECO'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
                eco_reports[brand] = eco_report

        return sales_report, sub_company_reports, eco_reports

    except Exception as e:
        logger.error(f"Error processing {company} file: {e}")
        raise ValueError(f"Error processing the {company} file: {e}")

def process_excel(file_path, company):
    if company in ['Canon', 'Canon Eldoret']:
        return process_canon(file_path, company)
    elif company in ['Jumra', 'Jumra Eldoret']:
        return process_jumra(file_path, company)
    else:
        raise ValueError(f"Invalid company: {company}")

def create_consolidated_excel(company, reports_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if company in ['Canon', 'Canon Eldoret']:
            sales_report, brand_reports, sku_eco_reports = reports_data
            logger.debug(f"Writing General Sales for {company}")
            sales_report.to_excel(writer, sheet_name='General Sales', index=False)
            for brand, report in brand_reports.items():
                if report is not None and not report.empty:
                    safe_sheet_name = f"{brand} ECO"[:31]
                    logger.debug(f"Writing {safe_sheet_name}")
                    report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            for sku, report in sku_eco_reports.items():
                if report is not None and not report.empty:
                    safe_sheet_name = f"{sku} ECO"[:31]
                    logger.debug(f"Writing {safe_sheet_name}")
                    report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
        elif company in ['Jumra', 'Jumra Eldoret']:
            sales_report, sub_company_reports, eco_reports = reports_data
            logger.debug(f"Writing General Sales for {company}")
            sales_report.to_excel(writer, sheet_name='General Sales', index=False)
            for sub_company, report in sub_company_reports.items():
                if report is not None and not report.empty:
                    safe_sheet_name = f"{sub_company} Report"[:31]
                    logger.debug(f"Writing {safe_sheet_name}")
                    report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            for brand, report in eco_reports.items():
                if report is not None and not report.empty:
                    safe_sheet_name = f"{brand} ECO"[:31]
                    logger.debug(f"Writing {safe_sheet_name}")
                    report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
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
            today_date = datetime.today().strftime('%Y-%m-%d')
            if company in ['Canon', 'Canon Eldoret']:
                sales_report, brand_reports, sku_eco_reports = reports_data
                # Generate HTML table without percentage classes
                html_table = sales_report.to_html(classes="table table-striped table-bordered", index=False)
                return render_template(
                    "canon_report.html",
                    sales_report=html_table,
                    brand_reports=brand_reports,
                    sku_eco_reports=sku_eco_reports,
                    today_date=today_date,
                    company=company,
                    download_filename=f"{company}_consolidated_report.xlsx",
                    original_filename=file.filename
                )
            elif company in ['Jumra', 'Jumra Eldoret']:
                sales_report, sub_company_reports, eco_reports = reports_data
                # Generate HTML table without percentage classes
                html_table = sales_report.to_html(classes="table table-striped table-bordered", index=False)
                return render_template(
                    "jumra_report.html",
                    sales_report=html_table,
                    sub_company_reports=sub_company_reports,
                    eco_reports=eco_reports,
                    today_date=today_date,
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
            reports_data = process_excel(file_path, company)
            excel_file = create_consolidated_excel(company, reports_data)
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