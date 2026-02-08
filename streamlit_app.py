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
    if 'TV' in cat_upper or 'TELEVISION' in cat_upper:
        return 'TV'
    elif 'MICROWAVE' in cat_upper or 'OVEN' in cat_upper:
        return 'MICROWAVE OVEN'
    elif 'REFRIGERATOR' in cat_upper or 'FRIDGE' in cat_upper or 'REF' in cat_upper:
        return 'REFRIGERATOR'
    elif 'AC' in cat_upper or 'AIR CONDITIONER' in cat_upper or 'AIRCONDITIONER' in cat_upper:
        return 'AC'
    elif 'WASHING' in cat_upper or 'WASHER' in cat_upper or 'WM' in cat_upper:
        return 'WASHING MACHINE'
    elif 'SMALL APPLIANCE' in cat_upper or 'SA' in cat_upper:
        return 'SMALL APPLIANCE'
    
    return None  # Categories not in target list

def process_data(product_df, osg_df):
    """Process and merge data for report generation"""
    
    # Identify columns - flexible column name matching
    product_cols = {
        'rbm': None,
        'branch': None,
        'category': None,
        'sold_price': None
    }
    
    osg_cols = {
        'branch': None,
        'category': None,
        'sold_price': None
    }
    
    # Find product file columns
    for col in product_df.columns:
        col_upper = col.upper()
        if 'RBM' in col_upper:
            product_cols['rbm'] = col
        elif 'BRANCH' in col_upper:
            product_cols['branch'] = col
        elif 'CATEGORY' in col_upper or 'ITEM CATEGORY' in col_upper:
            product_cols['category'] = col
        elif 'TAXABLE VALUE' in col_upper:
            product_cols['sold_price'] = col
        elif 'SOLD PRICE' in col_upper or 'ITEM RATE' in col_upper:
            if product_cols['sold_price'] is None:
                product_cols['sold_price'] = col
    
    # Find OSG file columns
    for col in osg_df.columns:
        col_upper = col.upper()
        if 'STORE NAME' in col_upper or 'BRANCH' in col_upper:
            osg_cols['branch'] = col
        elif 'CATEGORY' in col_upper:
            osg_cols['category'] = col
        elif 'SOLD PRICE' in col_upper:
            osg_cols['sold_price'] = col
    
    # Verify all required columns are found
    missing_product = [k for k, v in product_cols.items() if v is None]
    missing_osg = [k for k, v in osg_cols.items() if v is None]
    
    if missing_product or missing_osg:
        st.error(f"Missing columns - Product: {missing_product}, OSG: {missing_osg}")
        return None
    
    # Extract relevant columns
    product_clean = product_df[[
        product_cols['rbm'],
        product_cols['branch'],
        product_cols['category'],
        product_cols['sold_price']
    ]].copy()
    
    product_clean.columns = ['RBM', 'Branch', 'Category', 'Product_Sold_Price']
    
    osg_clean = osg_df[[
        osg_cols['branch'],
        osg_cols['category'],
        osg_cols['sold_price']
    ]].copy()
    
    osg_clean.columns = ['Branch', 'Category', 'OSG_Sold_Price']
    
    # Normalize categories
    product_clean['Category_Normalized'] = product_clean['Category'].apply(normalize_category)
    osg_clean['Category_Normalized'] = osg_clean['Category'].apply(normalize_category)
    
    # Remove rows with categories not in target list
    product_clean = product_clean[product_clean['Category_Normalized'].notna()].copy()
    osg_clean = osg_clean[osg_clean['Category_Normalized'].notna()].copy()
    
    # Aggregate Product data
    product_agg = product_clean.groupby(['RBM', 'Branch', 'Category_Normalized']).agg({
        'Product_Sold_Price': 'sum'
    }).reset_index()
    product_agg.columns = ['RBM', 'Branch', 'Category', 'Product_Sold_Price']
    
    # Aggregate OSG data
    osg_agg = osg_clean.groupby(['Branch', 'Category_Normalized']).agg({
        'OSG_Sold_Price': 'sum'
    }).reset_index()
    osg_agg.columns = ['Branch', 'Category', 'OSG_Sold_Price']
    
    # Merge data
    merged = product_agg.merge(osg_agg, on=['Branch', 'Category'], how='left')
    merged['OSG_Sold_Price'] = merged['OSG_Sold_Price'].fillna(0)
    
    # Add Target %
    merged['Target_%'] = merged['Category'].map(TARGET_MAP)
    
    # Calculate Value Conversion (%)
    merged['Value_Conversion_%'] = np.where(
        merged['Product_Sold_Price'] > 0,
        (merged['OSG_Sold_Price'] / merged['Product_Sold_Price'] * 100).round(2),
        0
    )
    
    # Calculate Need to Achieve Target (Value)
    merged['Need_to_Achieve_Target'] = np.maximum(
        (merged['Product_Sold_Price'] * merged['Target_%'] / 100 - merged['OSG_Sold_Price']).round(2),
        0
    )
    
    # Reorder columns
    final_df = merged[[
        'RBM',
        'Branch',
        'Category',
        'Product_Sold_Price',
        'OSG_Sold_Price',
        'Value_Conversion_%',
        'Need_to_Achieve_Target',
        'Target_%'
    ]].copy()
    
    return final_df

def create_excel_report(df):
    """Create Excel report with RBM-wise sheets"""
    
    wb = Workbook()
    wb.remove(wb.active)
    
    # Header style
    header_style = {
        'font': Font(bold=True, color='FFFFFF', size=11),
        'fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }
    
    data_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Get unique RBMs
    rbms = sorted(df['RBM'].unique())
    
    for rbm in rbms:
        # Filter data for this RBM
        rbm_data = df[df['RBM'] == rbm].copy()
        rbm_data = rbm_data.drop('RBM', axis=1)  # Remove RBM column from sheet
        rbm_data = rbm_data.sort_values(['Branch', 'Category'])
        
        # Create sheet with RBM name (sanitize sheet name)
        sheet_name = str(rbm)[:31]  # Excel sheet name max 31 chars
        ws = wb.create_sheet(sheet_name)
        
        # Add title
        ws['A1'] = f'Target Achievement Report - {rbm}'
        ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
        ws.merge_cells('A1:H1')
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Headers
        headers = [
            'Branch',
            'Category',
            'Product Sold Price',
            'OSG Sold Price',
            'Value Conversion (%)',
            'Need to Achieve Target (Value)',
            'Target %'
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
                cell.alignment = Alignment(horizontal='left' if col_idx <= 2 else 'right', vertical='center')
                
                # Number formatting
                if col_idx in [3, 4, 6]:  # Product Price, OSG Price, Need to Achieve
                    cell.number_format = '₹#,##0.00'
                elif col_idx in [5, 7]:  # Value Conversion %, Target %
                    cell.number_format = '0.00"%"'
        
        # Column widths
        ws.column_dimensions['A'].width = 25  # Branch
        ws.column_dimensions['B'].width = 20  # Category
        ws.column_dimensions['C'].width = 20  # Product Sold Price
        ws.column_dimensions['D'].width = 18  # OSG Sold Price
        ws.column_dimensions['E'].width = 22  # Value Conversion %
        ws.column_dimensions['F'].width = 28  # Need to Achieve Target
        ws.column_dimensions['G'].width = 12  # Target %
        
        # Add summary at bottom
        last_row = len(rbm_data) + 4
        ws.cell(row=last_row + 1, column=1, value='TOTAL').font = Font(bold=True)
        
        # Total formulas
        ws.cell(row=last_row + 1, column=3, value=f'=SUM(C4:C{last_row})')
        ws.cell(row=last_row + 1, column=3).number_format = '₹#,##0.00'
        ws.cell(row=last_row + 1, column=3).font = Font(bold=True)
        
        ws.cell(row=last_row + 1, column=4, value=f'=SUM(D4:D{last_row})')
        ws.cell(row=last_row + 1, column=4).number_format = '₹#,##0.00'
        ws.cell(row=last_row + 1, column=4).font = Font(bold=True)
        
        ws.cell(row=last_row + 1, column=5, value=f'=IF(C{last_row + 1}>0,(D{last_row + 1}/C{last_row + 1})*100,0)')
        ws.cell(row=last_row + 1, column=5).number_format = '0.00"%"'
        ws.cell(row=last_row + 1, column=5).font = Font(bold=True)
        
        ws.cell(row=last_row + 1, column=6, value=f'=SUM(F4:F{last_row})')
        ws.cell(row=last_row + 1, column=6).number_format = '₹#,##0.00'
        ws.cell(row=last_row + 1, column=6).font = Font(bold=True)
    
    # Create Summary sheet
    ws_summary = wb.create_sheet('Summary', 0)
    ws_summary['A1'] = 'OSG TARGET ACHIEVEMENT SUMMARY'
    ws_summary['A1'].font = Font(bold=True, size=16, color='1F4E78')
    ws_summary.merge_cells('A1:G1')
    ws_summary['A1'].alignment = Alignment(horizontal='center')
    
    # RBM-wise summary
    summary_data = df.groupby('RBM').agg({
        'Product_Sold_Price': 'sum',
        'OSG_Sold_Price': 'sum',
        'Need_to_Achieve_Target': 'sum'
    }).reset_index()
    
    summary_data['Value_Conversion_%'] = (
        summary_data['OSG_Sold_Price'] / summary_data['Product_Sold_Price'] * 100
    ).round(2)
    
    # Summary headers
    summary_headers = ['RBM', 'Product Sold Price', 'OSG Sold Price', 'Value Conversion (%)', 'Need to Achieve Target']
    for col_idx, header in enumerate(summary_headers, start=1):
        cell = ws_summary.cell(row=3, column=col_idx, value=header)
        cell.font = header_style['font']
        cell.fill = header_style['fill']
        cell.alignment = header_style['alignment']
        cell.border = header_style['border']
    
    # Summary data
    for row_idx, row_data in enumerate(summary_data.values, start=4):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
            cell.border = data_border
            cell.alignment = Alignment(horizontal='left' if col_idx == 1 else 'right', vertical='center')
            
            if col_idx in [2, 3, 5]:
                cell.number_format = '₹#,##0.00'
            elif col_idx == 4:
                cell.number_format = '0.00"%"'
    
    for col in range(1, 6):
        ws_summary.column_dimensions[get_column_letter(col)].width = 25
    
    return wb

def main():
    
    st.sidebar.header("📁 Upload Files")
    
    product_file = st.sidebar.file_uploader("Upload Product File (Excel)", type=['xlsx', 'xls'], key='product')
    osg_file = st.sidebar.file_uploader("Upload OSG File (Excel)", type=['xlsx', 'xls'], key='osg')
    
    if product_file and osg_file:
        
        with st.spinner("Loading files..."):
            product_df = pd.read_excel(product_file)
            osg_df = pd.read_excel(osg_file)
        
        st.success(f"✅ Files loaded! Product: {len(product_df)} rows, OSG: {len(osg_df)} rows")
        
        # Show column detection
        with st.expander("📋 Detected Columns"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Product File Columns:**")
                st.write(product_df.columns.tolist())
            with col2:
                st.write("**OSG File Columns:**")
                st.write(osg_df.columns.tolist())
        
        if st.button("🚀 Generate Report", type="primary"):
            
            with st.spinner("Processing data..."):
                processed_df = process_data(product_df, osg_df)
            
            if processed_df is not None:
                
                st.success(f"✅ Data processed successfully! {len(processed_df)} records")
                
                # Preview
                st.subheader("📊 Data Preview")
                st.dataframe(processed_df.head(20), use_container_width=True)
                
                # Statistics
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Product Sales", f"₹{processed_df['Product_Sold_Price'].sum():,.2f}")
                
                with col2:
                    st.metric("Total OSG Sales", f"₹{processed_df['OSG_Sold_Price'].sum():,.2f}")
                
                with col3:
                    overall_conversion = (processed_df['OSG_Sold_Price'].sum() / 
                                        processed_df['Product_Sold_Price'].sum() * 100)
                    st.metric("Overall Conversion", f"{overall_conversion:.2f}%")
                
                with col4:
                    st.metric("Total Gap", f"₹{processed_df['Need_to_Achieve_Target'].sum():,.2f}")
                
                # Generate Excel
                with st.spinner("Creating Excel report..."):
                    wb = create_excel_report(processed_df)
                    
                    # Save to BytesIO
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
                
                # Show RBM list
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
        - TV: **5%**
        - Microwave Oven: **5%**
        - Refrigerator: **2%**
        - AC: **1%**
        - Washing Machine: **3%**
        - Small Appliance: **2%**
        
        ### 📊 Report Structure:
        - **One sheet per RBM** with RBM name as sheet name
        - **Summary sheet** with overall RBM-wise statistics
        - Each sheet contains Branch-wise and Category-wise breakdown
        
        ### 📈 Columns in Report:
        1. Branch
        2. Category
        3. Product Sold Price
        4. OSG Sold Price
        5. Value Conversion (%)
        6. Need to Achieve Target (Value)
        7. Target %
        """)

if __name__ == "__main__":
    main()
