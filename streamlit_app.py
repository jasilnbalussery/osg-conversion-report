import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OSG Target Achievement Report", layout="wide")

st.title("📊 OSG Target Achievement Report Generator")

# Target percentages
TARGET_MAP = {
    'TV': 5,
    'MICROWAVE OVEN': 5,
    'REFRIGERATOR': 2,
    'AC': 1,
    'WASHING MACHINE': 3,
    'SMALL APPLIANCE': 2
}

def normalize_category(category):
    """Normalize category names to match target categories"""
    if pd.isna(category):
        return None

    cat_upper = str(category).upper().strip()

    # Exact matches first
    for target_cat in TARGET_MAP.keys():
        if cat_upper == target_cat:
            return target_cat

    # Partial matches
    if 'WASHING MACHINE' in cat_upper or 'WASHING' in cat_upper or 'WASHER' in cat_upper:
        return 'WASHING MACHINE'
    elif 'WM' == cat_upper or cat_upper.startswith('WM ') or ' WM' in cat_upper or cat_upper.endswith(' WM'):
        return 'WASHING MACHINE'
    elif 'TELEVISION' in cat_upper or cat_upper == 'TV' or cat_upper.startswith('TV ') or ' TV' in cat_upper or cat_upper.endswith(' TV'):
        return 'TV'
    elif 'MICROWAVE' in cat_upper or ('MICRO' in cat_upper and 'WAVE' in cat_upper):
        return 'MICROWAVE OVEN'
    elif 'OVEN' in cat_upper and 'MICROWAVE' not in cat_upper:
        return 'MICROWAVE OVEN'
    elif 'REFRIGERATOR' in cat_upper or 'FRIDGE' in cat_upper or 'REFG' in cat_upper:
        return 'REFRIGERATOR'
    elif 'REF' == cat_upper or cat_upper.startswith('REF ') or ' REF' in cat_upper or cat_upper.endswith(' REF'):
        return 'REFRIGERATOR'
    elif 'AIR CONDITIONER' in cat_upper or 'AIRCONDITIONER' in cat_upper or 'AIR CONDITION' in cat_upper:
        return 'AC'
    elif 'AC' == cat_upper or cat_upper.startswith('AC ') or ' AC' in cat_upper or cat_upper.endswith(' AC'):
        return 'AC'
    elif 'SMALL APPLIANCE' in cat_upper or 'SMALL APP' in cat_upper:
        return 'SMALL APPLIANCE'
    elif 'SA' == cat_upper or cat_upper.startswith('SA ') or ' SA' in cat_upper or cat_upper.endswith(' SA'):
        return 'SMALL APPLIANCE'

    return None


def process_data(product_df, osg_df):
    """Process and merge data for report generation"""

    product_cols = {'rbm': None, 'branch': None, 'category': None, 'sold_price': None}
    osg_cols = {'branch': None, 'category': None, 'sold_price': None}

    for col in product_df.columns:
        col_upper = col.upper()
        if 'RBM' in col_upper:
            product_cols['rbm'] = col
        elif 'BRANCH' in col_upper:
            product_cols['branch'] = col
        elif 'CATEGORY' in col_upper:
            product_cols['category'] = col
        elif 'SOLD PRICE' in col_upper:
            product_cols['sold_price'] = col
        elif 'TAXABLE VALUE' in col_upper and product_cols['sold_price'] is None:
            product_cols['sold_price'] = col

    for col in osg_df.columns:
        col_upper = col.upper()
        if 'STORE NAME' in col_upper or 'BRANCH' in col_upper:
            osg_cols['branch'] = col
        elif 'CATEGORY' in col_upper:
            osg_cols['category'] = col
        elif 'SOLD PRICE' in col_upper:
            osg_cols['sold_price'] = col

    missing_product = [k for k, v in product_cols.items() if v is None]
    missing_osg = [k for k, v in osg_cols.items() if v is None]

    if missing_product or missing_osg:
        st.error(f"Missing columns - Product: {missing_product}, OSG: {missing_osg}")
        return None

    product_clean = product_df[[
        product_cols['rbm'], product_cols['branch'],
        product_cols['category'], product_cols['sold_price']
    ]].copy()
    product_clean.columns = ['RBM', 'Branch', 'Category', 'Product_Sold_Price']

    osg_clean = osg_df[[
        osg_cols['branch'], osg_cols['category'], osg_cols['sold_price']
    ]].copy()
    osg_clean.columns = ['Branch', 'Category', 'OSG_Sold_Price']

    product_clean['Category_Normalized'] = product_clean['Category'].apply(normalize_category)
    osg_clean['Category_Normalized'] = osg_clean['Category'].apply(normalize_category)

    product_clean = product_clean[product_clean['Category_Normalized'].notna()].copy()
    osg_clean = osg_clean[osg_clean['Category_Normalized'].notna()].copy()

    product_agg = product_clean.groupby(['RBM', 'Branch', 'Category_Normalized']).agg({
        'Product_Sold_Price': 'sum'
    }).reset_index()
    product_agg.columns = ['RBM', 'Branch', 'Category', 'Product_Sold_Price']

    osg_agg = osg_clean.groupby(['Branch', 'Category_Normalized']).agg({
        'OSG_Sold_Price': 'sum'
    }).reset_index()
    osg_agg.columns = ['Branch', 'Category', 'OSG_Sold_Price']

    merged = product_agg.merge(osg_agg, on=['Branch', 'Category'], how='left')
    merged['OSG_Sold_Price'] = merged['OSG_Sold_Price'].fillna(0)

    # Target %
    merged['Target_%'] = merged['Category'].map(TARGET_MAP)

    # Value Conversion (%)
    merged['Value_Conversion_%'] = np.where(
        merged['Product_Sold_Price'] > 0,
        (merged['OSG_Sold_Price'] / merged['Product_Sold_Price'] * 100).round(2),
        0
    )

    # ── CHANGED: Need to Achieve is now a CONVERSION GAP (%), not a value ──
    merged['Need_to_Achieve_Target_%'] = np.maximum(
        (merged['Target_%'] - merged['Value_Conversion_%']).round(2),
        0
    )

    # Achievement status helper
    merged['Target_Achieved'] = merged['Value_Conversion_%'] >= merged['Target_%']

    final_df = merged[[
        'RBM',
        'Branch',
        'Category',
        'Product_Sold_Price',
        'OSG_Sold_Price',
        'Value_Conversion_%',
        'Target_%',
        'Need_to_Achieve_Target_%',
        'Target_Achieved'
    ]].copy()

    return final_df


def create_excel_report(df):
    """Create Excel report with RBM-wise sheets"""

    wb = Workbook()
    wb.remove(wb.active)

    header_style = {
        'font': Font(bold=True, color='FFFFFF', size=11),
        'fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
    }

    data_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # ── Conditional fills ──
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    green_font = Font(color='006100')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    red_font = Font(color='9C0006')

    rbms = sorted(df['RBM'].unique())

    for rbm in rbms:
        rbm_data = df[df['RBM'] == rbm].copy()
        rbm_data = rbm_data.drop(['RBM', 'Target_Achieved'], axis=1)
        rbm_data = rbm_data.sort_values(['Branch', 'Category'])

        sheet_name = str(rbm)[:31]
        ws = wb.create_sheet(sheet_name)

        # Title
        ws['A1'] = f'Target Achievement Report - {rbm}'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:G1')
        ws['A1'].alignment = Alignment(horizontal='center')

        # ── Updated headers ──
        headers = [
            'Branch',
            'Category',
            'Product Sold Price',
            'OSG Sold Price',
            'Value Conversion (%)',
            'Target %',
            'Need to Achieve Target (%)'
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = header_style['font']
            cell.fill = header_style['fill']
            cell.alignment = header_style['alignment']
            cell.border = header_style['border']

        # Data rows
        for row_idx, row_data in enumerate(rbm_data.values, start=4):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = data_border
                cell.alignment = Alignment(
                    horizontal='left' if col_idx <= 2 else 'right',
                    vertical='center'
                )

                # Number formatting
                if col_idx in [3, 4]:          # Product Price, OSG Price
                    cell.number_format = '₹#,##0.00'
                elif col_idx in [5, 6, 7]:     # All three % columns
                    cell.number_format = '0.00"%"'

            # ── Colour the "Need to Achieve" cell (col 7) ──
            need_cell = ws.cell(row=row_idx, column=7)
            conv_cell_val = row_data[4]   # Value_Conversion_%
            target_cell_val = row_data[5] # Target_%

            if conv_cell_val >= target_cell_val:
                need_cell.fill = green_fill
                need_cell.font = green_font
            else:
                need_cell.fill = red_fill
                need_cell.font = red_font

        # Column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 22
        ws.column_dimensions['F'].width = 14
        ws.column_dimensions['G'].width = 28

        # ── Summary row ──
        last_row = len(rbm_data) + 3
        summary_row = last_row + 2

        ws.cell(row=summary_row, column=1, value='TOTAL').font = Font(bold=True)

        # Total Product Sold Price
        ws.cell(row=summary_row, column=3, value=f'=SUM(C4:C{last_row})')
        ws.cell(row=summary_row, column=3).number_format = '₹#,##0.00'
        ws.cell(row=summary_row, column=3).font = Font(bold=True)

        # Total OSG Sold Price
        ws.cell(row=summary_row, column=4, value=f'=SUM(D4:D{last_row})')
        ws.cell(row=summary_row, column=4).number_format = '₹#,##0.00'
        ws.cell(row=summary_row, column=4).font = Font(bold=True)

        # Overall Conversion %
        ws.cell(row=summary_row, column=5,
                value=f'=IF(C{summary_row}>0,(D{summary_row}/C{summary_row})*100,0)')
        ws.cell(row=summary_row, column=5).number_format = '0.00"%"'
        ws.cell(row=summary_row, column=5).font = Font(bold=True)

    # ────────────── Summary sheet ──────────────
    ws_summary = wb.create_sheet('Summary', 0)
    ws_summary['A1'] = 'OSG TARGET ACHIEVEMENT SUMMARY'
    ws_summary['A1'].font = Font(bold=True, size=16, color='1F4E78')
    ws_summary.merge_cells('A1:F1')
    ws_summary['A1'].alignment = Alignment(horizontal='center')

    summary_data = df.groupby('RBM').agg({
        'Product_Sold_Price': 'sum',
        'OSG_Sold_Price': 'sum'
    }).reset_index()

    summary_data['Value_Conversion_%'] = np.where(
        summary_data['Product_Sold_Price'] > 0,
        (summary_data['OSG_Sold_Price'] / summary_data['Product_Sold_Price'] * 100).round(2),
        0
    )

    # Weighted-average target for each RBM
    rbm_target = df.groupby('RBM').apply(
        lambda g: np.average(g['Target_%'], weights=g['Product_Sold_Price'])
        if g['Product_Sold_Price'].sum() > 0 else 0
    ).reset_index(name='Weighted_Target_%')

    summary_data = summary_data.merge(rbm_target, on='RBM')
    summary_data['Need_to_Achieve_%'] = np.maximum(
        (summary_data['Weighted_Target_%'] - summary_data['Value_Conversion_%']).round(2), 0
    )

    summary_headers = [
        'RBM', 'Product Sold Price', 'OSG Sold Price',
        'Value Conversion (%)', 'Wtd. Target (%)', 'Need to Achieve (%)'
    ]
    for col_idx, header in enumerate(summary_headers, start=1):
        cell = ws_summary.cell(row=3, column=col_idx, value=header)
        cell.font = header_style['font']
        cell.fill = header_style['fill']
        cell.alignment = header_style['alignment']
        cell.border = header_style['border']

    for row_idx, row_data in enumerate(summary_data.values, start=4):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
            cell.border = data_border
            cell.alignment = Alignment(
                horizontal='left' if col_idx == 1 else 'right',
                vertical='center'
            )
            if col_idx in [2, 3]:
                cell.number_format = '₹#,##0.00'
            elif col_idx in [4, 5, 6]:
                cell.number_format = '0.00"%"'

        # Colour Need-to-Achieve column
        need_cell = ws_summary.cell(row=row_idx, column=6)
        if row_data[3] >= row_data[4]:   # conversion >= target
            need_cell.fill = green_fill
            need_cell.font = green_font
        else:
            need_cell.fill = red_fill
            need_cell.font = red_font

    for col in range(1, 7):
        ws_summary.column_dimensions[get_column_letter(col)].width = 25

    return wb


def main():

    st.sidebar.header("📁 Upload Files")

    product_file = st.sidebar.file_uploader(
        "Upload Product File (Excel)", type=['xlsx', 'xls'], key='product')
    osg_file = st.sidebar.file_uploader(
        "Upload OSG File (Excel)", type=['xlsx', 'xls'], key='osg')

    if product_file and osg_file:

        with st.spinner("Loading files..."):
            product_df = pd.read_excel(product_file)
            osg_df = pd.read_excel(osg_file)

        st.success(f"✅ Files loaded! Product: {len(product_df)} rows, OSG: {len(osg_df)} rows")

        with st.expander("📋 Detected Columns"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Product File Columns:**")
                st.write(product_df.columns.tolist())
            with col2:
                st.write("**OSG File Columns:**")
                st.write(osg_df.columns.tolist())

        with st.expander("🔍 Category Analysis (Debug Info)"):
            col1, col2 = st.columns(2)

            product_cat_col = None
            osg_cat_col = None

            for col in product_df.columns:
                if 'CATEGORY' in col.upper():
                    product_cat_col = col
                    break
            for col in osg_df.columns:
                if 'CATEGORY' in col.upper():
                    osg_cat_col = col
                    break

            with col1:
                st.write("**Product File Unique Categories:**")
                if product_cat_col:
                    st.dataframe(product_df[product_cat_col].value_counts().head(20),
                                 use_container_width=True)
                    sample_cats = product_df[product_cat_col].unique()[:10]
                    st.write("**Category Mapping Preview:**")
                    st.dataframe(pd.DataFrame({
                        'Original': sample_cats,
                        'Normalized': [normalize_category(c) for c in sample_cats]
                    }), use_container_width=True)

            with col2:
                st.write("**OSG File Unique Categories:**")
                if osg_cat_col:
                    st.dataframe(osg_df[osg_cat_col].value_counts().head(20),
                                 use_container_width=True)
                    sample_cats = osg_df[osg_cat_col].unique()[:10]
                    st.write("**Category Mapping Preview:**")
                    st.dataframe(pd.DataFrame({
                        'Original': sample_cats,
                        'Normalized': [normalize_category(c) for c in sample_cats]
                    }), use_container_width=True)

        if st.button("🚀 Generate Report", type="primary"):

            with st.spinner("Processing data..."):
                processed_df = process_data(product_df, osg_df)

            if processed_df is not None:

                st.success(f"✅ Data processed successfully! {len(processed_df)} records")

                st.subheader("📊 Category Distribution")
                st.dataframe(processed_df['Category'].value_counts(),
                             use_container_width=True)

                st.subheader("📋 Data Preview")
                st.dataframe(processed_df.head(20), use_container_width=True)

                # ── Updated metrics ──
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    st.metric("Total Product Sales",
                              f"₹{processed_df['Product_Sold_Price'].sum():,.2f}")
                with col2:
                    st.metric("Total OSG Sales",
                              f"₹{processed_df['OSG_Sold_Price'].sum():,.2f}")
                with col3:
                    overall_conv = (processed_df['OSG_Sold_Price'].sum() /
                                    processed_df['Product_Sold_Price'].sum() * 100)
                    st.metric("Overall Conversion", f"{overall_conv:.2f}%")
                with col4:
                    achieved = processed_df['Target_Achieved'].sum()
                    total = len(processed_df)
                    st.metric("Targets Achieved",
                              f"{achieved}/{total} ({achieved/total*100:.1f}%)")

                # Generate Excel
                with st.spinner("Creating Excel report..."):
                    wb = create_excel_report(processed_df)
                    buffer = BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)

                st.success("✅ Report generated successfully!")

                st.download_button(
                    label="📥 Download Excel Report",
                    data=buffer,
                    file_name="OSG_Target_Achievement_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.subheader("📑 Report Contains Following Sheets:")
                rbms = sorted(processed_df['RBM'].unique())
                st.write(f"**Summary Sheet** + **{len(rbms)} RBM Sheets:**")
                st.write(", ".join(rbms))

    else:
        st.info("👆 Please upload both Product and OSG files to generate the report")

        st.markdown("""
        ### 📋 Instructions:

        **Product File** should contain:
        - RBM
        - Branch
        - Category (or Item Category)
        - Sold Price (or Taxable Value)

        **OSG File** should contain:
        - Branch (or Store Name)
        - Category
        - Sold Price

        ### 🎯 Target Percentages:
        | Category | Target |
        |----------|--------|
        | TV | **5%** |
        | Microwave Oven | **5%** |
        | Refrigerator | **2%** |
        | AC | **1%** |
        | Washing Machine | **3%** |
        | Small Appliance | **2%** |

        ### 📊 Report Columns:
        1. Branch
        2. Category
        3. Product Sold Price
        4. OSG Sold Price
        5. **Value Conversion (%)** — actual OSG/Product %
        6. **Target %** — required conversion
        7. **Need to Achieve Target (%)** — gap = Target% − Conversion%
           - 🟢 Green = target already met
           - 🔴 Red = still needs improvement
        """)


if __name__ == "__main__":
    main()
