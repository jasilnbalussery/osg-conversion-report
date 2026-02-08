import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Set page config
st.set_page_config(page_title="OSG Target Achievement Analysis", layout="wide", initial_sidebar_state="expanded")

# Title
st.title("📊 OSG Target Achievement Analysis - RBM & Store Wise")

# Define target percentages
TARGET_PERCENTAGES = {
    'TV': 5,
    'AC': 1,
    'REF': 2,
    'WM': 3,
    'OVEN': 5,
    'SA': 2,
    'OTHER': 0  # No specific target mentioned for others
}

def normalize_category(category):
    """Normalize category names to match target categories"""
    if pd.isna(category):
        return 'OTHER'
    
    category_upper = str(category).upper().strip()
    
    # Map to standard categories
    if 'TV' in category_upper or 'TELEVISION' in category_upper:
        return 'TV'
    elif 'AC' in category_upper or 'AIR CONDITIONER' in category_upper or 'AIRCONDITIONER' in category_upper:
        return 'AC'
    elif 'REF' in category_upper or 'REFRIGERATOR' in category_upper or 'FRIDGE' in category_upper:
        return 'REF'
    elif 'WM' in category_upper or 'WASHING' in category_upper or 'WASHER' in category_upper:
        return 'WM'
    elif 'OVEN' in category_upper or 'MICROWAVE' in category_upper:
        return 'OVEN'
    elif 'SA' in category_upper or 'SMALL APPLIANCE' in category_upper:
        return 'SA'
    else:
        return 'OTHER'

@st.cache_data
def load_data(osg_file, product_file):
    """Load and process the uploaded files"""
    try:
        # Load OSG data
        osg_df = pd.read_excel(osg_file)
        
        # Load Product data
        product_df = pd.read_excel(product_file)
        
        return osg_df, product_df
    except Exception as e:
        st.error(f"Error loading files: {str(e)}")
        return None, None

def process_data(osg_df, product_df):
    """Process and merge data for analysis"""
    
    # Normalize category in OSG data
    osg_df['Normalized_Category'] = osg_df['Category'].apply(normalize_category)
    
    # Normalize category in Product data
    product_df['Normalized_Category'] = product_df['Item Category'].apply(normalize_category)
    
    # Convert date columns
    if 'Date' in osg_df.columns:
        osg_df['Date'] = pd.to_datetime(osg_df['Date'], errors='coerce')
    
    if 'Date' in product_df.columns:
        product_df['Date'] = pd.to_datetime(product_df['Date'], errors='coerce')
    
    # Calculate OSG sales by Store, RBM, and Category
    osg_summary = osg_df.groupby(['Store Name', 'Normalized_Category']).agg({
        'Invoice No': 'count',
        'Sold Price': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    osg_summary.columns = ['Store Name', 'Category', 'OSG_Invoice_Count', 'OSG_Sales_Value', 'OSG_Quantity']
    
    # Calculate Product sales by Store (Branch), RBM, and Category
    product_summary = product_df.groupby(['Branch', 'RBM', 'Normalized_Category']).agg({
        'Invoice Number': 'count',
        'Taxable Value': 'sum',
        'QTY': 'sum'
    }).reset_index()
    product_summary.columns = ['Store Name', 'RBM', 'Category', 'Product_Invoice_Count', 'Product_Sales_Value', 'Product_Quantity']
    
    # Merge OSG and Product data
    merged_df = product_summary.merge(
        osg_summary, 
        on=['Store Name', 'Category'], 
        how='left'
    )
    
    # Fill NaN values with 0
    merged_df['OSG_Invoice_Count'] = merged_df['OSG_Invoice_Count'].fillna(0)
    merged_df['OSG_Sales_Value'] = merged_df['OSG_Sales_Value'].fillna(0)
    merged_df['OSG_Quantity'] = merged_df['OSG_Quantity'].fillna(0)
    
    # Calculate target values
    merged_df['Target_Percentage'] = merged_df['Category'].map(TARGET_PERCENTAGES)
    merged_df['Target_Value'] = (merged_df['Product_Sales_Value'] * merged_df['Target_Percentage'] / 100).round(2)
    
    # Calculate achievement
    merged_df['Achievement_Value'] = merged_df['OSG_Sales_Value']
    merged_df['Achievement_Percentage'] = np.where(
        merged_df['Target_Value'] > 0,
        (merged_df['Achievement_Value'] / merged_df['Target_Value'] * 100).round(2),
        0
    )
    
    # Calculate gap
    merged_df['Gap_Value'] = (merged_df['Target_Value'] - merged_df['Achievement_Value']).round(2)
    merged_df['Gap_Percentage'] = (100 - merged_df['Achievement_Percentage']).round(2)
    
    return merged_df, osg_df, product_df

def create_summary_metrics(df):
    """Create summary metrics"""
    
    total_target = df['Target_Value'].sum()
    total_achievement = df['Achievement_Value'].sum()
    overall_achievement_pct = (total_achievement / total_target * 100) if total_target > 0 else 0
    total_gap = total_target - total_achievement
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Target Value", f"₹{total_target:,.2f}")
    
    with col2:
        st.metric("Total Achievement", f"₹{total_achievement:,.2f}")
    
    with col3:
        st.metric("Achievement %", f"{overall_achievement_pct:.2f}%")
    
    with col4:
        st.metric("Gap to Target", f"₹{total_gap:,.2f}", delta=f"-{(total_gap/total_target*100):.1f}%" if total_target > 0 else "0%")

def main():
    
    # File uploaders
    st.sidebar.header("📁 Upload Files")
    
    osg_file = st.sidebar.file_uploader("Upload OSG Data (Excel)", type=['xlsx', 'xls'], key='osg')
    product_file = st.sidebar.file_uploader("Upload Product Data (Excel)", type=['xlsx', 'xls'], key='product')
    
    if osg_file and product_file:
        
        # Load data
        with st.spinner("Loading data..."):
            osg_df, product_df = load_data(osg_file, product_file)
        
        if osg_df is not None and product_df is not None:
            
            # Process data
            with st.spinner("Processing data..."):
                merged_df, osg_raw, product_raw = process_data(osg_df, product_df)
            
            st.success(f"✅ Data loaded successfully! OSG Records: {len(osg_raw)}, Product Records: {len(product_raw)}")
            
            # Filters
            st.sidebar.header("🔍 Filters")
            
            # RBM Filter
            rbm_list = ['All'] + sorted(merged_df['RBM'].dropna().unique().tolist())
            selected_rbm = st.sidebar.selectbox("Select RBM", rbm_list)
            
            # Store Filter
            if selected_rbm != 'All':
                store_list = ['All'] + sorted(merged_df[merged_df['RBM'] == selected_rbm]['Store Name'].dropna().unique().tolist())
            else:
                store_list = ['All'] + sorted(merged_df['Store Name'].dropna().unique().tolist())
            selected_store = st.sidebar.selectbox("Select Store", store_list)
            
            # Category Filter
            category_list = ['All'] + sorted(merged_df['Category'].unique().tolist())
            selected_category = st.sidebar.selectbox("Select Category", category_list)
            
            # Apply filters
            filtered_df = merged_df.copy()
            
            if selected_rbm != 'All':
                filtered_df = filtered_df[filtered_df['RBM'] == selected_rbm]
            
            if selected_store != 'All':
                filtered_df = filtered_df[filtered_df['Store Name'] == selected_store]
            
            if selected_category != 'All':
                filtered_df = filtered_df[filtered_df['Category'] == selected_category]
            
            # Display summary metrics
            st.header("📈 Overall Performance")
            create_summary_metrics(filtered_df)
            
            # Tabs for different views
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["🎯 RBM-wise Report", "🏪 Store-wise Report", "📊 Category-wise Report", "📉 Detailed Data", "📋 Raw Data"])
            
            with tab1:
                st.subheader("RBM-wise Target Achievement")
                
                rbm_summary = filtered_df.groupby('RBM').agg({
                    'Product_Sales_Value': 'sum',
                    'Target_Value': 'sum',
                    'Achievement_Value': 'sum',
                    'Gap_Value': 'sum'
                }).reset_index()
                
                rbm_summary['Achievement_Percentage'] = np.where(
                    rbm_summary['Target_Value'] > 0,
                    (rbm_summary['Achievement_Value'] / rbm_summary['Target_Value'] * 100).round(2),
                    0
                )
                
                rbm_summary = rbm_summary.sort_values('Achievement_Percentage', ascending=False)
                
                # Display table
                st.dataframe(
                    rbm_summary.style.format({
                        'Product_Sales_Value': '₹{:,.2f}',
                        'Target_Value': '₹{:,.2f}',
                        'Achievement_Value': '₹{:,.2f}',
                        'Gap_Value': '₹{:,.2f}',
                        'Achievement_Percentage': '{:.2f}%'
                    }).background_gradient(subset=['Achievement_Percentage'], cmap='RdYlGn', vmin=0, vmax=100),
                    use_container_width=True
                )
                
                # RBM Chart
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    name='Target Value',
                    x=rbm_summary['RBM'],
                    y=rbm_summary['Target_Value'],
                    marker_color='lightblue'
                ))
                fig.add_trace(go.Bar(
                    name='Achievement Value',
                    x=rbm_summary['RBM'],
                    y=rbm_summary['Achievement_Value'],
                    marker_color='green'
                ))
                fig.update_layout(
                    title='RBM-wise Target vs Achievement',
                    xaxis_title='RBM',
                    yaxis_title='Value (₹)',
                    barmode='group',
                    height=500
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                st.subheader("Store-wise Target Achievement")
                
                store_summary = filtered_df.groupby(['RBM', 'Store Name']).agg({
                    'Product_Sales_Value': 'sum',
                    'Target_Value': 'sum',
                    'Achievement_Value': 'sum',
                    'Gap_Value': 'sum'
                }).reset_index()
                
                store_summary['Achievement_Percentage'] = np.where(
                    store_summary['Target_Value'] > 0,
                    (store_summary['Achievement_Value'] / store_summary['Target_Value'] * 100).round(2),
                    0
                )
                
                store_summary = store_summary.sort_values('Achievement_Percentage', ascending=False)
                
                # Display table
                st.dataframe(
                    store_summary.style.format({
                        'Product_Sales_Value': '₹{:,.2f}',
                        'Target_Value': '₹{:,.2f}',
                        'Achievement_Value': '₹{:,.2f}',
                        'Gap_Value': '₹{:,.2f}',
                        'Achievement_Percentage': '{:.2f}%'
                    }).background_gradient(subset=['Achievement_Percentage'], cmap='RdYlGn', vmin=0, vmax=100),
                    use_container_width=True
                )
                
                # Top 10 and Bottom 10 stores
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### 🏆 Top 10 Stores")
                    top_10 = store_summary.head(10)
                    fig_top = px.bar(
                        top_10,
                        x='Store Name',
                        y='Achievement_Percentage',
                        color='Achievement_Percentage',
                        color_continuous_scale='Greens',
                        title='Top 10 Performing Stores'
                    )
                    fig_top.update_layout(height=400)
                    st.plotly_chart(fig_top, use_container_width=True)
                
                with col2:
                    st.markdown("#### 📉 Bottom 10 Stores")
                    bottom_10 = store_summary.tail(10)
                    fig_bottom = px.bar(
                        bottom_10,
                        x='Store Name',
                        y='Achievement_Percentage',
                        color='Achievement_Percentage',
                        color_continuous_scale='Reds',
                        title='Bottom 10 Performing Stores'
                    )
                    fig_bottom.update_layout(height=400)
                    st.plotly_chart(fig_bottom, use_container_width=True)
            
            with tab3:
                st.subheader("Category-wise Target Achievement")
                
                category_summary = filtered_df.groupby('Category').agg({
                    'Product_Sales_Value': 'sum',
                    'Target_Percentage': 'first',
                    'Target_Value': 'sum',
                    'Achievement_Value': 'sum',
                    'Gap_Value': 'sum'
                }).reset_index()
                
                category_summary['Achievement_Percentage'] = np.where(
                    category_summary['Target_Value'] > 0,
                    (category_summary['Achievement_Value'] / category_summary['Target_Value'] * 100).round(2),
                    0
                )
                
                category_summary = category_summary.sort_values('Achievement_Percentage', ascending=False)
                
                # Display table
                st.dataframe(
                    category_summary.style.format({
                        'Product_Sales_Value': '₹{:,.2f}',
                        'Target_Percentage': '{:.0f}%',
                        'Target_Value': '₹{:,.2f}',
                        'Achievement_Value': '₹{:,.2f}',
                        'Gap_Value': '₹{:,.2f}',
                        'Achievement_Percentage': '{:.2f}%'
                    }).background_gradient(subset=['Achievement_Percentage'], cmap='RdYlGn', vmin=0, vmax=100),
                    use_container_width=True
                )
                
                # Category pie chart
                fig_pie = px.pie(
                    category_summary,
                    values='Achievement_Value',
                    names='Category',
                    title='Category-wise OSG Sales Distribution'
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
                # Category bar chart
                fig_cat = go.Figure()
                fig_cat.add_trace(go.Bar(
                    name='Target Value',
                    x=category_summary['Category'],
                    y=category_summary['Target_Value'],
                    marker_color='lightblue'
                ))
                fig_cat.add_trace(go.Bar(
                    name='Achievement Value',
                    x=category_summary['Category'],
                    y=category_summary['Achievement_Value'],
                    marker_color='green'
                ))
                fig_cat.update_layout(
                    title='Category-wise Target vs Achievement',
                    xaxis_title='Category',
                    yaxis_title='Value (₹)',
                    barmode='group',
                    height=500
                )
                st.plotly_chart(fig_cat, use_container_width=True)
            
            with tab4:
                st.subheader("Detailed RBM → Store → Category Report")
                
                # Create detailed view
                detailed_view = filtered_df.copy()
                detailed_view = detailed_view.sort_values(['RBM', 'Store Name', 'Category'])
                
                # Display with formatting
                st.dataframe(
                    detailed_view[['RBM', 'Store Name', 'Category', 'Product_Sales_Value', 
                                  'Target_Percentage', 'Target_Value', 'Achievement_Value', 
                                  'Achievement_Percentage', 'Gap_Value', 'Gap_Percentage']].style.format({
                        'Product_Sales_Value': '₹{:,.2f}',
                        'Target_Percentage': '{:.0f}%',
                        'Target_Value': '₹{:,.2f}',
                        'Achievement_Value': '₹{:,.2f}',
                        'Achievement_Percentage': '{:.2f}%',
                        'Gap_Value': '₹{:,.2f}',
                        'Gap_Percentage': '{:.2f}%'
                    }).background_gradient(subset=['Achievement_Percentage'], cmap='RdYlGn', vmin=0, vmax=100),
                    use_container_width=True,
                    height=600
                )
                
                # Download button
                csv = detailed_view.to_csv(index=False)
                st.download_button(
                    label="📥 Download Detailed Report as CSV",
                    data=csv,
                    file_name=f"osg_target_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with tab5:
                st.subheader("Raw Data Explorer")
                
                data_choice = st.radio("Select Data to View", ["OSG Data", "Product Data"])
                
                if data_choice == "OSG Data":
                    st.dataframe(osg_raw, use_container_width=True, height=600)
                    st.info(f"Total Records: {len(osg_raw)}")
                else:
                    st.dataframe(product_raw, use_container_width=True, height=600)
                    st.info(f"Total Records: {len(product_raw)}")
    
    else:
        st.info("👆 Please upload both OSG Data and Product Data files to begin analysis.")
        
        # Show instructions
        st.markdown("""
        ### 📋 Instructions:
        
        1. **Upload OSG Data File** - Should contain columns like:
           - Store Name, Category, Invoice No, Sold Price, Quantity, etc.
        
        2. **Upload Product Data File** - Should contain columns like:
           - Branch, RBM, BDM, Item Category, Invoice Number, Taxable Value, QTY, etc.
        
        3. **Target Percentages Applied**:
           - TV: 5%
           - AC: 1%
           - REF: 2%
           - WM: 3%
           - OVEN: 5%
           - SA: 2%
           - OTHER: 0%
        
        4. The app will automatically:
           - Match categories between OSG and Product data
           - Calculate targets based on product sales
           - Show achievement and gaps for each RBM, Store, and Category
        """)

if __name__ == "__main__":
    main()
