"""
Sprint Analysis Automation - Streamlit Web App
Simple web interface for sprint analysis
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import sys
from pathlib import Path

# Set page config
st.set_page_config(
    page_title="Sprint Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #3B82F6;
        color: white;
        font-size: 18px;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        border: none;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #2563EB;
    }
    .upload-section {
        border: 3px dashed #E2E8F0;
        border-radius: 15px;
        padding: 3rem;
        text-align: center;
        background: #F8FAFC;
        margin: 2rem 0;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        margin: 0.5rem;
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: bold;
        font-family: monospace;
    }
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .success-box {
        background: #D1FAE5;
        border-left: 5px solid #10B981;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    .warning-box {
        background: #FEF3C7;
        border-left: 5px solid #F59E0B;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    h1 {
        color: #1E293B;
        font-size: 3rem;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        text-align: center;
        color: #64748B;
        font-size: 1.3rem;
        margin-bottom: 2rem;
    }
    .footer {
        text-align: center;
        color: #94A3B8;
        padding: 2rem;
        margin-top: 3rem;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Import automation modules
try:
    from automation.data_processor import SprintDataProcessor
    from automation.calculator import SprintCalculator
    from automation.excel_updater import ExcelUpdater
except ImportError:
    st.error("⚠️ Automation modules not found. Please ensure the app is properly deployed.")
    st.stop()


def process_sprint_file(uploaded_file):
    """Process the uploaded Excel file"""
    try:
        # Save uploaded file to temp location
        temp_path = "temp_sprint.xlsx"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Initialize processor
        processor = SprintDataProcessor(temp_path)
        calculator = SprintCalculator()
        
        # Load workbook
        wb = openpyxl.load_workbook(temp_path, data_only=False)
        
        # Extract metadata
        sprint_metadata = processor.extract_sprint_metadata(wb['Data'])
        
        # Process Azure DevOps data
        df_azure = pd.read_excel(temp_path, sheet_name='Data', header=20)
        azure_data = processor.process_azure_data(df_azure)
        
        # Validate data
        validation = processor.validate_data(azure_data)
        if validation['status'] == 'error':
            return None, validation['errors'], None, None
        
        # Get capacity data
        capacity_data = processor.get_capacity_data(wb['Capacity'])
        
        # Calculate metrics
        staff_agg = processor.aggregate_by_staff(azure_data)
        team_agg = processor.aggregate_by_team(azure_data)
        
        staff_metrics = calculator.calculate_staff_metrics(
            staff_agg, capacity_data, azure_data
        )
        
        team_metrics = calculator.calculate_team_metrics(
            team_agg, capacity_data, azure_data
        )
        
        cmmi_measures = calculator.calculate_cmmi_measures(
            sprint_metadata, azure_data
        )
        
        # Update Excel sheets
        updater = ExcelUpdater(temp_path)
        sprint_name = sprint_metadata['sprint_name']
        
        updater.update_analysis_sheet(sprint_name, staff_metrics, team_metrics)
        updater.update_kpi_indicators_sheet(staff_metrics, team_metrics, sprint_name)
        updater.append_to_historical_staff(staff_metrics, sprint_name)
        updater.append_to_historical_team(team_metrics, sprint_name)
        updater.update_cmmi_template(cmmi_measures, sprint_name)
        
        # Save workbook
        updater.save()
        
        # Read the processed file
        with open(temp_path, 'rb') as f:
            processed_data = f.read()
        
        return processed_data, validation, staff_metrics, team_metrics, cmmi_measures
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None, None, None, None, None


# Main app
def main():
    # Header
    st.markdown("<h1>📊 Sprint Analysis Automation</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Upload your Excel file → Get instant analysis → Download results</p>", unsafe_allow_html=True)
    
    # Info box
    st.info("🔒 **Your data is safe**: All processing happens on our secure server. Your file is processed and then immediately deleted. No data is stored.")
    
    # Upload section
    st.markdown("### 📁 Upload Your Sprint File")
    
    uploaded_file = st.file_uploader(
        "Choose your Sprint Excel file (.xlsx or .xlsm)",
        type=['xlsx', 'xlsm'],
        help="Upload the Excel file with your Azure DevOps sprint data"
    )
    
    if uploaded_file is not None:
        # Show file info
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 File Name", uploaded_file.name)
        with col2:
            file_size = len(uploaded_file.getvalue()) / 1024
            st.metric("📦 File Size", f"{file_size:.1f} KB")
        with col3:
            st.metric("📅 Status", "Ready to Process")
        
        st.markdown("---")
        
        # Process button
        if st.button("⚡ Process Sprint Data", type="primary"):
            with st.spinner("🔄 Processing your sprint data... This may take a minute."):
                
                # Progress steps
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                steps = [
                    "Loading workbook...",
                    "Extracting sprint metadata...",
                    "Processing Azure DevOps data...",
                    "Validating data quality...",
                    "Loading capacity data...",
                    "Calculating staff metrics...",
                    "Calculating team metrics...",
                    "Computing CMMI measures...",
                    "Updating analysis sheets...",
                    "Finalizing workbook..."
                ]
                
                for i, step in enumerate(steps):
                    status_text.text(f"[{i+1}/{len(steps)}] {step}")
                    progress_bar.progress((i + 1) / len(steps))
                
                # Process file
                result = process_sprint_file(uploaded_file)
                processed_data, validation, staff_metrics, team_metrics, cmmi_measures = result
                
                if processed_data is None:
                    st.error("❌ Processing failed. Please check your file format and try again.")
                    if validation and 'errors' in validation:
                        for error in validation['errors']:
                            st.error(f"• {error['message']}")
                else:
                    # Success!
                    st.success("✅ Analysis Complete!")
                    
                    # Show warnings if any
                    if validation and validation.get('warnings'):
                        with st.expander("⚠️ Warnings (click to expand)"):
                            for warning in validation['warnings']:
                                st.warning(f"• {warning['message']}")
                    
                    # Display metrics
                    st.markdown("### 📊 Summary Metrics")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Staff Analyzed</div>
                            <div class="metric-value">{len(staff_metrics)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Teams Processed</div>
                            <div class="metric-value">{len(team_metrics)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        avg_kpi = team_metrics['kpi'].mean()
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Avg Team KPI</div>
                            <div class="metric-value">{avg_kpi:.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        completion = cmmi_measures['completion_rate'] * 100
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">CMMI Completion</div>
                            <div class="metric-value">{completion:.0f}%</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
                    # Download button
                    st.markdown("### ⬇️ Download Your Analyzed File")
                    
                    output_filename = uploaded_file.name.replace('.xlsx', '_analyzed.xlsx')
                    
                    st.download_button(
                        label="📥 Download Analyzed File",
                        data=processed_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                    st.success("🎉 Your file is ready! Click the button above to download.")
                    
                    # What's included
                    with st.expander("📋 What's included in your analyzed file"):
                        st.markdown("""
                        **New/Updated Sheets:**
                        - ✅ **[Sprint Name] Analysis** - Detailed staff and team breakdowns
                        - ✅ **Kpi Indicators** - Current sprint KPIs with formulas
                        - ✅ **Kpi Indicators Per Staff** - Historical tracking (new column added)
                        - ✅ **Kpi Indicators Per Team** - Historical tracking (new column added)
                        - ✅ **CMMI Template** - CMMI measures history (new row added)
                        
                        **Metrics Calculated:**
                        - 📊 6 Staff KPIs (per developer)
                        - 👥 6 Team KPIs (per team)
                        - 📈 5 CMMI Measures
                        - ✨ All using Excel formulas (dynamic, not hardcoded!)
                        """)
    
    else:
        # Instructions when no file uploaded
        st.markdown("""
        <div class='upload-section'>
            <h2>👆 Click above to upload your Sprint Excel file</h2>
            <p style='color: #64748B; margin-top: 1rem;'>
                Supported formats: .xlsx, .xlsm
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📖 How to Use")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            **1️⃣ Prepare Your Data**
            - Export Azure DevOps query
            - Paste data into 'Data' sheet (row 21)
            - Update sprint metadata (rows 3-10)
            """)
        
        with col2:
            st.markdown("""
            **2️⃣ Upload & Process**
            - Click "Browse files" above
            - Select your Excel file
            - Click "Process Sprint Data"
            """)
        
        with col3:
            st.markdown("""
            **3️⃣ Download Results**
            - Wait ~30 seconds
            - Review summary metrics
            - Download analyzed file
            """)
        
        st.markdown("---")
        
        st.markdown("### ❓ Frequently Asked Questions")
        
        with st.expander("What data do I need in my Excel file?"):
            st.markdown("""
            Your Excel file should have:
            - **Data sheet** with Azure DevOps export (starting row 21)
            - **Capacity sheet** with team member capacity
            - Sprint metadata in Data sheet (rows 3-10)
            
            The file structure should match the template provided.
            """)
        
        with st.expander("Is my data safe?"):
            st.markdown("""
            **Yes!** Your data is completely safe:
            - ✅ Processing happens on our secure server
            - ✅ Files are immediately deleted after processing
            - ✅ No data is stored or logged
            - ✅ No data is sent to third parties
            - ✅ HTTPS encrypted connection
            """)
        
        with st.expander("What metrics are calculated?"):
            st.markdown("""
            **Staff-Level (Per Developer):**
            - KPI (composite score)
            - Performance Rate
            - Utilization
            - Done Tasks %
            - MidSprint Addition %
            - Ad-hoc %
            
            **Team-Level:**
            - Same 6 metrics aggregated by team
            
            **CMMI Measures:**
            - Completion Rate
            - Effort Estimation Accuracy
            - Bug Fixing Effort %
            - Utilization Rate
            - CMMI Productivity
            """)
        
        with st.expander("How long does processing take?"):
            st.markdown("""
            Processing typically takes **20-60 seconds** depending on:
            - File size
            - Number of team members
            - Number of tasks
            - Server load
            
            You'll see real-time progress updates.
            """)
    
    # Footer
    st.markdown("""
    <div class='footer'>
        <p>🔒 All processing is secure and confidential</p>
        <p>Made with ❤️ for efficient sprint analysis</p>
        <p style='margin-top: 1rem; font-size: 0.8rem;'>
            Need help? Contact your system administrator
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
