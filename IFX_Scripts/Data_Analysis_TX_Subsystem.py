#
#
# ----------------- Test Data Analysis Script (Flat Version) -----------------
#
import pandas as pd
import os
import re
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
# import numpy as np


# Suppress warnings from openpyxl for a cleaner output
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 1. User Configuration ----------------------------------------------------------------------------------------------------------------------------------
# <<<<<<<<<<<<<<< PLEASE CONFIGURE THESE PATHS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# DATA_FOLDER = 'C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/FE_Test_CS/'
# OUTPUT_FOLDER = 'C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/FE_Test_CS/Data_Analysis'
DATA_FOLDER = 'C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/BE_Test_CS'
OUTPUT_FOLDER = 'C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/BE_Test_CS'
OUTPUT_FOLDER_PLOTS_TX_POWER = 'C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/BE_Test_CS/Plots_TX_Power'
# <<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

DATA_ANALYSIS_EXCEL_REPORT_NAME = 'Yield_Cpk_Correlation_report.xlsx'
PDF_REPORT_NAME = 'tx_power_overview.pdf'
#HEATMAPS_FOLDER = 'correlation_heatmaps'
PERFORM_YIELD_CPK_ANALYSES = True
PERFORM_TX_CORRELATION_ANALYSES = True
PERFORM_TX_POWER_PLOTTING = True

# --- 2. Test Module Definitions -----------------------------------------------------------------------------------------------------------------------------
TEST_MODULES = {
    'TXGE': (50000, 50999), 'TXVC': (51000, 51999),
    'DPLL': (52000, 52999), 'TXPA': (53000, 53999),
    'TXPB': (54000, 54999), 'TXPC': (55000, 55999),
    'TXPD': (56000, 56999), 'TXLO': (57000, 57999),
    'TXPS': (58000, 58999),
}
TX_MODULE_NAMES = list(TEST_MODULES.keys())


# --- 3. Main Execution Block --------------------------------------------------------------------------------------------------------------------------------
print("\n")
print("----------------------------------------------- Starting Test Data Analysis -----------------------------------------------")

# Setup directories and paths
if not os.path.exists(DATA_FOLDER) or not os.listdir(DATA_FOLDER):
    os.makedirs(DATA_FOLDER, exist_ok=True)
    print(f"ERROR: The '{
          DATA_FOLDER}' is empty. Please add your CSV files and run again.")
else:
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    #heatmap_dir = os.path.join(OUTPUT_FOLDER, HEATMAPS_FOLDER)
    #os.makedirs(heatmap_dir, exist_ok=True)

    excel_path = os.path.join(OUTPUT_FOLDER, DATA_ANALYSIS_EXCEL_REPORT_NAME)
    pdf_path = os.path.join(OUTPUT_FOLDER, PDF_REPORT_NAME)
    os.makedirs(OUTPUT_FOLDER_PLOTS_TX_POWER, exist_ok=True)

    csv_files = [f for f in os.listdir(
        DATA_FOLDER) if f.lower().endswith('.csv')]
    # This list will collect data from all files for the final PDF plot
    all_tx_power_data = []

    # Open the Excel writer to create one report file with multiple sheets
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:  # openpyxl xlsxwriter
        # Loop through each CSV file in the data folder
        for filename in csv_files:
            print(f"\nProcessing {filename}...")
            file_path = os.path.join(DATA_FOLDER, filename)
            # Excel sheet names have a 31-character limit
            sheet_name = os.path.splitext(filename)[0][:31]

            ##
            # ----------------- START: Data Loading and Cleaning -----------------
            ##
            # Find the header row where the main table starts
            header_row = None
            try:
                with open(file_path, 'r', encoding='latin1') as f:
                    for i, line in enumerate(f):
                        if 'Test Nr' in line:
                            header_row = i
                            break
            except Exception as e:
                print(f"  - Could not read {file_path}: {e}")

            if header_row is None:
                print("  - Warning: Could not find header 'Test Nr'. Skipping file.")
                continue

            # Load the data starting from the header row
            df = pd.read_csv(file_path, sep=';', low_memory=False)
            df.columns = df.columns.str.strip()

            # Find where the raw device data columns start
            first_raw_col_index = next(
                (i for i, col in enumerate(df.columns) if col.strip().isdigit()), -1)

            if first_raw_col_index == -1:
                print("  - Warning: Could not identify raw data columns. Skipping file.")
                continue

            # Separate summary statistics from raw data
            summary_df = df.iloc[:, :first_raw_col_index].copy()
            raw_df = df.iloc[:, first_raw_col_index:].copy()
            raw_data_df = df.drop(df.columns[2:13], axis=1)
            raw_data_df.index = raw_data_df['Test Nr']

            # Clean the summary data
            for col in ['Test Nr', 'Cpk', 'Yield']:
                if col in summary_df.columns:
                    summary_df[col] = pd.to_numeric(
                        summary_df[col], errors='coerce')
            summary_df.dropna(subset=['Test Nr'], inplace=True)
            summary_df['Test Nr'] = summary_df['Test Nr'].astype(int)
            meta_cols = ['Test Nr', 'Test Name',
                         'Unit', 'Mean', 'Cpk', 'Yield']
            df_meta = df[[col for col in meta_cols if col in df.columns]].copy()
            for col in ['Test Nr', 'Yield', 'Cpk', 'Mean']:
                df_meta[col] = pd.to_numeric(df_meta[col], errors='coerce')
            df_meta.dropna(subset=['Test Nr', 'Yield', 'Cpk'], inplace=True)
            df_meta['Test Nr'] = df_meta['Test Nr'].astype(int)

            # Transpose and clean the raw data
            raw_df.index = summary_df['Test Nr']
            raw_df_transposed = raw_df.transpose()
            raw_df_transposed = raw_df_transposed.apply(
                pd.to_numeric, errors='coerce')
            ##
            # ----------------- END: Data Loading and Cleaning -----------------
            ##

            ##
            # ----------------- START: Yield and Cpk Analysis -----------------
            ##
            if PERFORM_YIELD_CPK_ANALYSES:
                print("  - Analyzing Yield and Cpk...")
                yield_reports = []
                cpk_reports = []
                cpk_reports2 = []
                for name, (start, end) in TEST_MODULES.items():
                    module_df = df_meta[df_meta['Test Nr'].between(start, end)]
    
                    # Yield < 100%, sorted ascending
                    yr = module_df[module_df['Yield'] < 100][[
                        'Test Name', 'Test Nr', 'Yield']].sort_values(by='Yield')
                    if not yr.empty:
                        yield_reports.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Yield < 100% ----------------"}]), yr]))
                    else:
                        yield_reports.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Yield < 100% ----------------"}])]))
    
                    # Cpk < 1.67, sorted ascending
                    cr = module_df[module_df['Cpk'] < 1.67][[
                        'Test Name', 'Test Nr', 'Cpk']].sort_values(by='Cpk')
                    if not cr.empty:
                        cpk_reports.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Cpk < 1.67 ----------------"}]), cr]))
                    else:
                        cpk_reports.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Cpk < 1.67 ----------------"}])]))
                        
                    # Cpk > 20, sorted ascending
                    cr2 = module_df[module_df['Cpk'] > 20][[
                        'Test Name', 'Test Nr', 'Cpk']].sort_values(by='Cpk')
                    if not cr2.empty:
                        cpk_reports2.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Cpk > 20 ----------------"}]), cr2]))
                    else:
                        cpk_reports2.append(pd.concat([pd.DataFrame(
                            [{'Test Name': f"---------------- {name} Cpk > 20 ----------------"}])]))
    
                # Write yield/cpk reports to Excel
                pd.concat(yield_reports, ignore_index=True).to_excel(excel_writer=writer, sheet_name=sheet_name, index=False, startrow=0)
                pd.concat(cpk_reports, ignore_index=True).to_excel(excel_writer=writer, sheet_name=sheet_name, index=False, startcol=4)
                pd.concat(cpk_reports2, ignore_index=True).to_excel(excel_writer=writer, sheet_name=sheet_name, index=False, startcol=8)
                print(f"  - Yield/Cpk analysis written to sheet: {sheet_name}")
            ##
            # ----------------- END: Yield and Cpk Analysis -----------------
            ##

            ##
            # ----------------- START: Correlation Analysis -----------------
            ##
            if PERFORM_TX_CORRELATION_ANALYSES:
                print("  - Analyzing correlations...")
                # Assign a module name to each test number
                module_list = []
                for test_nr in summary_df['Test Nr']:
                    module_name = 'UNKNOWN'
                    # Find which module range the test number falls into
                    for name, (start, end) in TEST_MODULES.items():
                        if start <= int(test_nr) <= end:
                            module_name = name
                            break
                    module_list.append(module_name)
                summary_df['Module'] = module_list
                raw_data_df['Module'] = module_list
    
                # Get test numbers for all TX modules
                tx_test_nrs = summary_df[summary_df['Module'].isin(TX_MODULE_NAMES)]['Test Nr']
    
                # Calculate the full correlation matrix
                corr_matrix = raw_df_transposed.corr(method='spearman', min_periods=1, numeric_only=True)
                #corr_matrix = raw_df_transposed.corr(method='pearson', min_periods=1, numeric_only=True)
    
                # Find correlation pairs involving at least one TX test
                pairs = []
                for test1 in tx_test_nrs:
                    if test1 not in corr_matrix: continue
                    corrs = corr_matrix[test1]
                    # Filter for correlations between 0.6 and 1.0
                    valid_corrs = corrs[(corrs >= 0.6) & (corrs < 1.0)]
                    for test2, val in valid_corrs.items():
                        if test1 < test2: # Avoid duplicate pairs (e.g., A-B and B-A)
                            pairs.append((test1, test2, val))
    
                # Write correlation results below the Cpk results
                next_col = writer.sheets[sheet_name].max_column + 2
                next_row = 0
                pd.DataFrame(["Correlation Analysis (absolute Coeff. >= 0.9)"]).to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=next_row, startcol=next_col, index=False, header=False)
    
                if pairs:
                    # Create a DataFrame with the results and add test names
                    corr_df = pd.DataFrame(pairs, columns=['Test Nr 1', 'Test Nr 2', 'Correlation'])
                    test_info = summary_df[['Test Nr', 'Test Name']].set_index('Test Nr')
                    corr_df = corr_df.join(test_info.rename(columns={'Test Name': 'Test Name 1'}), on='Test Nr 1')
                    corr_df = corr_df.join(test_info.rename(columns={'Test Name': 'Test Name 2'}), on='Test Nr 2')
                    corr_df = corr_df[['Test Name 1', 'Test Nr 1', 'Test Name 2', 'Test Nr 2', 'Correlation']].sort_values('Test Nr 1', ascending=True)
    
                    strong = corr_df[(corr_df['Correlation'] >= 0.9) | (corr_df['Correlation'] <= -0.9)]
                    moderate = corr_df[(corr_df['Correlation'] >= 0.6) & (corr_df['Correlation'] < 0.8)]
    
                    # Write "Strong" correlations to Excel
                    current_row = next_row + 1
                    pd.DataFrame(["Strong correlation (absolute Coeff. 0.9 to 1.0)"]).to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=current_row, startcol=next_col, index=False, header=False)
                    if not strong.empty:
                        strong.to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=current_row + 1, startcol=next_col, index=False)
                        current_row += len(strong) + 3
                    else:
                        pd.DataFrame(["None"]).to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=current_row + 1, startcol=next_col, index=False, header=False)
                        current_row += 3
    
                    # Write "Moderate" correlations to Excel
                    # pd.DataFrame(["Moderate correlation (0.6 to 0.8)"]).to_excel(writer, sheet_name, startrow=current_row, index=False, header=False)
                    # if not moderate.empty:
                    #     moderate.to_excel(writer, sheet_name, startrow=current_row + 1, index=False)
                    # else:
                    #     pd.DataFrame(["None"]).to_excel(writer, sheet_name, startrow=current_row + 1, index=False, header=False)
                else:
                      pd.DataFrame(["No correlations found for the TX tests."]).to_excel(excel_writer=writer, sheet_name=sheet_name, startrow=next_row+2, startcol=next_col, index=False, header=False)
                  
                # Generate heatmaps for each TX module
                # for module in TX_MODULE_NAMES:
                #     module_tests = summary_df[summary_df['Module'] == module]['Test Nr']
                #     module_tests_in_raw = list(set(module_tests) & set(raw_df_transposed.columns))
                #     if len(module_tests_in_raw) < 2: continue
    
                #     module_corr = raw_df_transposed[module_tests_in_raw].corr().abs()
                #     strong_pairs = module_corr[(module_corr >= 0.8) & (module_corr < 1.0)].stack().index
                #     #if not strong_pairs.any(): continue
    
                #     strong_tests = sorted(list(set([i for t in strong_pairs for i in t])))
                #     if len(strong_tests) < 2: continue
    
                #     plt.figure(figsize=(12, 10))
                #     sns.heatmap(raw_df_transposed[strong_tests].corr(), annot=True, cmap='viridis', fmt=".2f")
                #     plt.title(f'Strong Correlation Heatmap for {module} ({sheet_name})')
                #     plt.savefig(os.path.join(heatmap_dir, f"{sheet_name}_{module}_heatmap.png"), bbox_inches='tight')
                #     plt.close()
                ##
                # ----------------- END: Correlation Analysis -----------------
                ##
                # After processing all files, create the final reports
                print(f"\nYield, Cpk, and Correlation analyses saved to '{excel_path}'")
            
            #
            # ----------------- START: TX Power Data Extraction -----------------
            #
            if PERFORM_TX_POWER_PLOTTING:
                print("  - Extracting TX power data for plotting...")
                # Determine temperature from filename
                temp = 'Unknown'
                if 'S1' in filename or 'B1' in filename:
                    temp = 'Hot (135°C)'
                elif 'S2' in filename:
                    temp = 'Cold (-40°C)'
                elif 'S3' in filename or 'B2' in filename:
                    temp = 'Ambient (25°C)'
    
                if temp != 'Unknown':
                    # Filter for the relevant TXPA test modules, then remove the Module column
                    raw_data_df = raw_data_df[raw_data_df['Module'].isin(
                        ['TXPA', 'TXPB', 'TXPC'])]
                    raw_data_df = raw_data_df.drop('Module', axis=1)
                    for _, row in raw_data_df.iterrows():
                        test_name = row['Test Name']
                        test_nr = row['Test Nr']
                        # Parse metadata from the test name
                        channel_match = re.search(r'Tx([1-8])', test_name)
                        freq_match = re.search(r'(76|77|81)', test_name)
                        voltage_match = re.search(
                            r'(D095|D105)', test_name, re.IGNORECASE)
                        lut_match = re.search(
                            r'FwLu\d{1,3}', test_name, re.IGNORECASE)
                        is_valid = all([channel_match, freq_match, voltage_match, lut_match])
                        if is_valid and test_nr in raw_df_transposed:
                            v = voltage_match.group(1).upper()
                            voltage = 'VMIN' if v in ['D095'] else 'VMAX'
                            # Create a temporary DataFrame for this test's data
                            df_for_plot = pd.DataFrame(
                                {'Power_dBm': raw_df_transposed[test_nr].dropna()})
                            df_for_plot['Channel'] = f"TX{channel_match.group(1)}"
                            df_for_plot['Frequency'] = f"{freq_match.group(1)}GHz"
                            df_for_plot['Voltage'] = voltage
                            df_for_plot['lut'] = f"{lut_match.group(0)}"
                            df_for_plot['Temperature'] = temp
                            # Append to the master list for plotting later
                            all_tx_power_data.append(df_for_plot)
                    if not all_tx_power_data:
                        print("\nNo valid TX power data was found to generate the PDF report.")
                    else:
                        full_tx_df = pd.concat(all_tx_power_data, ignore_index=True,)
                        all_tx_power_data = []
                ##
                # ----------------- END: TX Power Data Extraction -----------------
                ##
        
                #
                # ----------------- START: Generate Combined TX Power Plot -----------------
                #
                # Create a separate plot for each unique LUT setting
                for lut in sorted(full_tx_df['lut'].unique()):
                    lut_data = full_tx_df[full_tx_df['lut'] == lut]
                    lut_data.index = range(len(lut_data.index))
                    lut_Digits = ''.join(char for char in lut if char.isdigit())
                    # Use seaborn's catplot to create a faceted grid of boxplots
                    sns.set_theme()
                    g = sns.relplot(
                        data=lut_data, x=range(len(lut_data.index)), y="Power_dBm", 
                        hue="Channel", size="Voltage", style="Frequency",
                        size_order=['VMAX', 'VMIN'], style_order=['76GHz', '77GHz', '81GHz'],
                        markers=['v', 'o', 'X'], dashes=False, legend="full",
                        kind="scatter")
                    g.set_titles(f"TX output power data for the LUT {lut_Digits}, T = {temp}")
                    g.set_axis_labels("", "TX power (dBm)")
                    plt.title(label=f"TX output power data for the LUT {lut_Digits}, T = {temp}")
                    #plt.show()
                    # Save the current figure to the PDF
                    plot_path = OUTPUT_FOLDER_PLOTS_TX_POWER + f"/{sheet_name}_TX_Power_LUT{lut_Digits}.png"
                    g.savefig(plot_path, format='png', dpi=600)
                    plt.close(g.fig)
                print(f"TX Power overview plot saved to '{OUTPUT_FOLDER_PLOTS_TX_POWER}'")
                ##
                # ----------------- END: Generate Combined TX Power Plot -----------------
                ##
    print("\n----------------------------------------------- Test Data Analysis Ending -----------------------------------------------")
