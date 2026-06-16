import sys
import io
import pandas as pd
import re
import xlsxwriter

import math
def get_axis_bounds(valid_vals):
    valid_floats = []
    for v in valid_vals:
        try: valid_floats.append(float(v))
        except: pass
    if not valid_floats: return {}
    v_min, v_max = min(valid_floats), max(valid_floats)
    v_span = v_max - v_min if v_max != v_min else (abs(v_min) * 0.1 if v_min != 0 else 1.0)
    return {'min': v_min - v_span * 0.1, 'max': v_max + v_span * 0.1}

def main():
    if len(sys.argv) < 2:
        print("Usage: python xlsx_generator.py <output_xlsx_path>")
        sys.exit(1)
        
    xlsx_path = sys.argv[1]
    
    # Read all CSV input from standard input
    raw_data = sys.stdin.read()
    if not raw_data.strip():
        print("Error: No data provided on stdin")
        sys.exit(1)
        
    try:
        # Split into lines to extract units
        raw_lines = [line.strip() for line in raw_data.split('\n') if line.strip() or ',' in line]
        headers = [h.strip() for h in raw_lines[0].split(',')]
        units = [u.strip() for u in raw_lines[1].split(',')]
        # Pad units if shorter
        units = units + [''] * (len(headers) - len(units))
    except Exception as e:
        print(f"Error parsing headers/units: {e}")
        sys.exit(1)
        
    try:
        # Read clean tabular data using pandas (skipping row 1 [units] and row 2 [empty])
        df = pd.read_csv(io.StringIO(raw_data), header=0, skiprows=[1, 2])
    except Exception as e:
        print(f"Error parsing CSV data: {e}")
        sys.exit(1)
        
    if df.empty or len(df.columns) < 2:
        print("Error: Tabular data is empty or missing columns")
        sys.exit(1)
        
    # Setup the Excel workbook
    workbook = xlsxwriter.Workbook(xlsx_path)
    
    # ── SHEET 1: SUMMARY ──
    summary_sheet = workbook.add_worksheet('Summary')
    summary_sheet.hide_gridlines(2) # show gridlines
    
    title_format = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#3B82F6', 'font_name': 'Segoe UI'})
    sub_format = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#F97316', 'font_name': 'Segoe UI'})
    bold_format = workbook.add_format({'bold': True, 'font_size': 11, 'font_name': 'Segoe UI'})
    text_format = workbook.add_format({'font_size': 11, 'font_name': 'Segoe UI'})
    
    summary_sheet.write('B2', 'ProMatX', title_format)
    summary_sheet.write('B3', 'Radiation Attenuation Calculator — Interactive Reports', sub_format)
    
    summary_sheet.write('B5', 'This workbook contains fully dynamic, editable calculation sheets and native charts:', bold_format)
    summary_sheet.write('B7', '1. Data Tab — Tabular numerical calculations of all parameters and compositions.', text_format)
    summary_sheet.write('B8', '2. Charts Tab — Fully editable comparison line charts dynamically linked to the Data tab.', text_format)
    summary_sheet.write('B10', '💡 TIP: Double-click any chart in Excel to customize colors, styles, titles, or axes!', bold_format)

    # Parse physical, mechanical, and optical parameters by composition
    props_list = [
        # Physical
        ("Molar Mass", "Physical"),
        ("Molar Volume", "Physical"),
        ("Avg Boron Distance", "Physical"),
        ("Oxygen Packing Density", "Physical"),
        ("Oxygen Molar Volume", "Physical"),
        ("Ion Concentration", "Physical"),
        ("Inter-ionic Distance", "Physical"),
        ("Polaron Radius", "Physical"),
        ("Field Strength", "Physical"),
        # Mechanical
        ("Packing Density", "Mechanical"),
        ("Dissociation Energy", "Mechanical"),
        ("Youngs Modulus", "Mechanical"),
        ("Bulk Modulus", "Mechanical"),
        ("Shear Modulus", "Mechanical"),
        ("Longitudinal Modulus", "Mechanical"),
        ("Poissons Ratio", "Mechanical"),
        ("Fractal Bond Connectivity", "Mechanical"),
        ("Vickers Hardness", "Mechanical"),
        # Optical
        ("Refractive Index", "Optical"),
        ("Dielectric Constant", "Optical"),
        ("Optical Dielectric Constant", "Optical"),
        ("Reflection Loss", "Optical"),
        ("Molar Refractivity", "Optical"),
        ("Metallization Criterion", "Optical"),
        ("Molar Polarizability", "Optical"),
        ("Electronic Polarizability", "Optical"),
        ("Transmission Coefficient", "Optical"),
        ("Optical Electronegativity", "Optical"),
        ("Linear Susceptibility", "Optical"),
        ("Non-linear Susceptibility", "Optical"),
        ("Non-linear Refractive Index", "Optical"),
        ("Density", "Physical"),
        ("CN", "Physical"),
        ("Bandgap", "Physical"),
        ("Dopant", "Physical")
    ]

    # Properties that are strings / non-numeric — skip individual charts for these
    STRING_ONLY_PROPS = {"CN", "Dopant"}

    # Display name overrides (internal name → chart label)
    DISPLAY_NAMES = {"Bandgap": "Bandgap Energy"}
    
    # Identify unique compositions and build mapping
    compositions = []
    comp_props = {}  # comp -> {base_param: (value, unit)}
    
    for col in df.columns:
        match = re.match(r"(.+?)\s*\[(.+?)\]", col)
        if match:
            base_param = match.group(1).strip()
            comp_name = match.group(2).strip()
        else:
            base_param = col
            comp_name = "Default"
            
        # Check if this is one of our parameters
        matched_param = None
        for p, cat in props_list:
            if p == base_param:
                matched_param = p
                break
        
        if matched_param:
            if comp_name not in comp_props:
                comp_props[comp_name] = {}
                compositions.append(comp_name)
            
            # Fetch value from df and unit from mapping
            val = df[col].iloc[0]
            unit = header_to_unit.get(col, "") if 'header_to_unit' in locals() else ""
            comp_props[comp_name][matched_param] = (val, unit)
                
    # If we found any properties, write them to the Summary sheet!
    if compositions:
        # Table styling formats
        header_format = workbook.add_format({
            'bold': True,
            'font_color': '#ffffff',
            'bg_color': '#3B82F6',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Segoe UI'
        })
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'right',
            'font_name': 'Segoe UI'
        })
        
        # Add summary table starting at row 12
        summary_sheet.write('B12', 'Physical, Mechanical, & Optical Properties Summary', bold_format)
        
        # Table Headers
        table_headers = ['Property', 'Unit'] + compositions
        for c_idx, h_text in enumerate(table_headers):
            summary_sheet.write(13, 1 + c_idx, h_text, header_format)
            
        current_row = 14
        
        # Print category by category
        for category in ["Physical", "Mechanical", "Optical"]:
            # Check if there's any property in this category that was actually calculated
            cat_props = [p for p, cat in props_list if cat == category]
            has_cat = False
            for comp in compositions:
                for cp in cat_props:
                    if cp in comp_props[comp]:
                        has_cat = True
                        break
            
            if has_cat:
                # Add category header row
                summary_sheet.write(current_row, 1, f"--- {category} Properties ---", bold_format)
                for c_idx in range(1, len(table_headers)):
                    summary_sheet.write(current_row, 1 + c_idx, "", bold_format)
                current_row += 1
                
                # Print each property in this category
                for cp in cat_props:
                    # Verify at least one composition calculated it
                    if not any(cp in comp_props[comp] for comp in compositions):
                        continue
                        
                    # Write property name
                    summary_sheet.write(current_row, 1, cp, text_format)
                    
                    # Fetch unit from the first composition that has it
                    unit = ""
                    for comp in compositions:
                        if cp in comp_props[comp]:
                            # Look up in units list
                            col_full_name = f"{cp} [{comp}]"
                            try:
                                col_idx = headers.index(col_full_name)
                                unit = units[col_idx]
                            except:
                                unit = ""
                            break
                    summary_sheet.write(current_row, 2, unit, cell_format)
                    
                    # Write values for each composition
                    for c_idx, comp in enumerate(compositions):
                        if cp in comp_props[comp]:
                            val = comp_props[comp][cp][0]
                            if pd.isna(val):
                                summary_sheet.write(current_row, 3 + c_idx, 'N/A', cell_format)
                            elif isinstance(val, (int, float)):
                                summary_sheet.write_number(current_row, 3 + c_idx, float(val), cell_format)
                            else:
                                summary_sheet.write_string(current_row, 3 + c_idx, str(val), cell_format)
                        else:
                            summary_sheet.write(current_row, 3 + c_idx, 'N/A', cell_format)
                    current_row += 1
                    
        # Auto-adjust column widths for the Summary sheet table
        summary_sheet.set_column(1, 1, 35) # Property column
        summary_sheet.set_column(2, 2, 18) # Unit column
        for c_idx in range(len(compositions)):
            summary_sheet.set_column(3 + c_idx, 3 + c_idx, 22) # Composition columns
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#ffffff',
        'bg_color': '#3B82F6',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_name': 'Segoe UI'
    })
    unit_format = workbook.add_format({
        'italic': True,
        'font_color': '#4b5563',
        'bg_color': '#f3f4f6',
        'border': 1,
        'align': 'center',
        'font_name': 'Segoe UI'
    })
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'right',
        'font_name': 'Segoe UI'
    })
    char_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'font_name': 'Segoe UI'
    })
    
    # 1. Separate columns into main (80 energies) and 25-energies
    STANDARD_25_ENERGIES = [
        0.015, 0.02, 0.03, 0.04, 0.05, 0.06, 0.08, 0.1, 0.15, 0.2, 
        0.3, 0.4, 0.5, 0.6, 0.8, 1.0, 1.5, 2.0, 3.0, 4.0, 
        5.0, 6.0, 8.0, 10.0, 15.0
    ]
    
    cols_25 = []
    cols_main = [df.columns[0]] # Energy column
    
    for col in df.columns[1:]:
        if col == "Isotope":
            cols_main.append(col)
            continue
            
        # Coerce column to numeric to properly identify NaN and count valid values
        numeric_col = pd.to_numeric(df[col], errors='coerce')
        valid_count = numeric_col.notna().sum()
        
        if valid_count <= 25 and len(df) > 25 and valid_count > 0:
            df[col] = numeric_col
            cols_25.append(col)
        else:
            cols_main.append(col)
            
    has_25 = len(cols_25) > 0
    
    df_main = df[cols_main]
    if has_25:
        df_25 = pd.DataFrame({'Energy': STANDARD_25_ENERGIES})
        for col in cols_25:
            # Map values to standard 25 energies using pandas Series mapping to align them correctly and avoid mismatches
            temp_series = pd.Series(df[col].values, index=df[df.columns[0]])
            df_25[col] = temp_series.reindex(STANDARD_25_ENERGIES).values
            
    header_to_unit = {h: u for h, u in zip(headers, units)}
    
    def write_data_sheet(sheet, dataframe, sheet_headers):
        sheet.hide_gridlines(2)
        for col_idx, col_name in enumerate(sheet_headers):
            sheet.write(0, col_idx, col_name, header_format)
            sheet.write(1, col_idx, header_to_unit.get(col_name, ''), unit_format)
            
        for row_idx, row in dataframe.iterrows():
            excel_row = int(row_idx) + 2
            for col_idx, col_name in enumerate(sheet_headers):
                val = row[col_name]
                if pd.isna(val):
                    sheet.write(excel_row, col_idx, '', cell_format)
                elif isinstance(val, (int, float)):
                    sheet.write_number(excel_row, col_idx, float(val), cell_format)
                else:
                    sheet.write_string(excel_row, col_idx, str(val), char_format)
                    
        for col_idx, col_name in enumerate(sheet_headers):
            max_len = max(len(col_name), 10)
            sheet.set_column(col_idx, col_idx, max_len + 3)

    # ── SHEET 2: GAMMA DATA ──
    # Filter out scalar property columns — only keep gamma-ray shielding columns
    scalar_props = {
        "Molar Mass", "Molar Volume", "Avg Boron Distance", "Oxygen Packing Density",
        "Oxygen Molar Volume", "Ion Concentration", "Inter-ionic Distance", "Polaron Radius",
        "Field Strength", "Packing Density", "Dissociation Energy", "Youngs Modulus",
        "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus", "Poissons Ratio",
        "Fractal Bond Connectivity", "Vickers Hardness",
        "Refractive Index", "Dielectric Constant", "Optical Dielectric Constant",
        "Reflection Loss", "Molar Refractivity", "Metallization Criterion", "Molar Polarizability",
        "Electronic Polarizability", "Transmission Coefficient", "Optical Electronegativity",
        "Linear Susceptibility", "Non-linear Susceptibility", "Non-linear Refractive Index",
        "Density", "CN", "Bandgap", "Dopant"
    }
    gamma_cols = [df_main.columns[0]]  # Energy column
    for col in df_main.columns[1:]:
        match_bp = re.match(r"(.+?)\s*\[(.+?)\]", col)
        base = match_bp.group(1).strip() if match_bp else col
        if base not in scalar_props:
            gamma_cols.append(col)
    df_gamma = df_main[gamma_cols]
    data_sheet = workbook.add_worksheet('Gamma Data')
    write_data_sheet(data_sheet, df_gamma, gamma_cols)
    
    # ── SHEET 3: 25 ENERGIES DATA (Conditional) ──
    if has_25:
        data_25_sheet = workbook.add_worksheet('25 Energies Data')
        cols_25_with_energy = ['Energy'] + cols_25
        write_data_sheet(data_25_sheet, df_25, cols_25_with_energy)
        
    # ── SHEET 4: CHARTS ──
    charts_sheet = workbook.add_worksheet('Charts')
    charts_sheet.hide_gridlines(2)
    
    groups = {}
    
    SKIP_CHARTS = {
        "Molar Mass", "Molar Volume", "Avg Boron Distance", "Oxygen Packing Density", 
        "Oxygen Molar Volume", "Ion Concentration", "Inter-ionic Distance", "Polaron Radius", 
        "Field Strength", "Packing Density", "Dissociation Energy", "Youngs Modulus", 
        "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus", "Poissons Ratio", 
        "Fractal Bond Connectivity", "Vickers Hardness", "Refractive Index", "Dielectric Constant", 
        "Optical Dielectric Constant",
        "Reflection Loss", "Molar Refractivity", "Metallization Criterion", "Molar Polarizability",
        "Electronic Polarizability", "Transmission Coefficient", "Optical Electronegativity",
        "Linear Susceptibility", "Non-linear Susceptibility", "Non-linear Refractive Index",
        "Density", "CN", "Bandgap", "Dopant"
    }
    
    def process_columns_for_groups(dataframe, sheet_name):
        for col_idx, col in enumerate(dataframe.columns):
            if col == "Energy" or col == "Isotope":
                continue
                
            match = re.match(r"(.+?)\s*\[(.+?)\]", col)
            if match:
                base_param = match.group(1).strip()
                comp_name = match.group(2).strip()
            else:
                base_param = col
                comp_name = "Default"
                
            if base_param in SKIP_CHARTS:
                continue
                
            if base_param not in groups:
                groups[base_param] = []
            groups[base_param].append({
                'comp_name': comp_name,
                'col_idx': col_idx,
                'sheet_name': sheet_name,
                'num_rows': len(dataframe)
            })

    process_columns_for_groups(df_gamma, 'Gamma Data')
    if has_25:
        process_columns_for_groups(df_25, '25 Energies Data')
        
    row_offset = 2
    
    # Refined scientific color palette (similar to matplotlib's default tab10 palette)
    palette = [
        '#3B82F6', '#F97316', '#2ca02c', '#d62728', 
        '#9467bd', '#8c564b', '#e377c2', '#7f7f7f'
    ]
    
    for base_param, members in groups.items():
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
        
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': '#cccccc', 'width': 1, 'dash_type': 'solid'}, 'fill': {'color': '#ffffff'}})
        
        for idx, member in enumerate(members):
            comp_name = member['comp_name']
            sheet_name = member['sheet_name']
            col_idx = member['col_idx']
            num_rows = member['num_rows']
            
            col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
            energy_col_letter = xlsxwriter.utility.xl_col_to_name(0)
            
            series_color = palette[idx % len(palette)]
            
            start_row = 3
            end_row = start_row + num_rows - 1
            
            sheet_ref = f"'{sheet_name}'" if " " in sheet_name else sheet_name
            
            chart.add_series({
                'name': comp_name if comp_name != "Default" else base_param,
                'categories': f"={sheet_ref}!${energy_col_letter}${start_row}:${energy_col_letter}${end_row}",
                'values': f"={sheet_ref}!${col_letter}${start_row}:${col_letter}${end_row}",
                'line': {'color': series_color, 'width': 1.5},
                'marker': {
                    'type': 'circle',
                    'size': 4.5,
                    'fill': {'color': series_color},
                    'border': {'color': series_color}
                }
            })
            
        chart.set_title({
            'name': f'{base_param} Comparison',
            'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
        })
        
        chart.set_legend({
            'position': 'bottom',
            'font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'}
        })
        
        x_axis_config = {
            'name': 'Energy (MeV)',
            'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'},
            'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}},
            'line': {'color': '#cccccc', 'width': 1},
            'crossing': 'min'
        }
        y_axis_config = {
            'name': base_param,
            'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
            'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'},
            'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}},
            'line': {'color': '#cccccc', 'width': 1},
            'crossing': 'min'
        }
        
        if any(x in base_param for x in ["MAC", "LAC", "MFP", "HVL", "TVL", "EBF", "EABF"]):
            x_axis_config['log_base'] = 10
            y_axis_config['log_base'] = 10
        else:
            x_axis_config['log_base'] = 10
            
        chart.set_x_axis(x_axis_config)
        chart.set_y_axis(y_axis_config)
        
        chart.set_size({'width': 850, 'height': 500})
        
        charts_sheet.insert_chart(f'B{row_offset}', chart)
        row_offset += 27
    
    # ── PROPERTY CHARTS SHEET (Physical / Mechanical / Optical) ──
    # Uses comp_props and compositions built earlier in the Summary section.
    # Each property gets a line chart with composition codes on X-axis.
    if compositions and comp_props:
        # Write a hidden data sheet for property charts
        prop_data_sheet = workbook.add_worksheet('PropData')
        prop_data_sheet.hide()
        
        prop_header_fmt = workbook.add_format({
            'bold': True, 'font_color': '#ffffff', 'bg_color': '#3B82F6',
            'border': 1, 'align': 'center', 'font_name': 'Segoe UI'
        })
        prop_cell_fmt = workbook.add_format({
            'border': 1, 'align': 'right', 'font_name': 'Segoe UI'
        })
        prop_label_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'font_name': 'Segoe UI',
            'bg_color': '#f3f4f6'
        })
        
        # Write header row: Property | S1 | S2 | ...
        prop_data_sheet.write(0, 0, 'Property', prop_header_fmt)
        for ci, comp in enumerate(compositions):
            prop_data_sheet.write(0, ci + 1, comp, prop_header_fmt)
        
        # Collect all properties that have at least one value
        all_calculated_props = []
        for p, cat in props_list:
            if any(p in comp_props.get(comp, {}) for comp in compositions):
                all_calculated_props.append((p, cat))
        
        # Write data rows: one row per property
        prop_row_map = {}  # prop_name -> excel_row (0-indexed)
        for row_i, (p, cat) in enumerate(all_calculated_props):
            excel_row = row_i + 1
            prop_row_map[p] = excel_row
            prop_data_sheet.write(excel_row, 0, p, prop_label_fmt)
            for ci, comp in enumerate(compositions):
                val = comp_props.get(comp, {}).get(p, (None,))[0]
                if val is not None and not pd.isna(val) and val != 'N/A':
                    try:
                        prop_data_sheet.write_number(excel_row, ci + 1, float(val), prop_cell_fmt)
                    except ValueError:
                        prop_data_sheet.write_string(excel_row, ci + 1, str(val), prop_cell_fmt)
                else:
                    prop_data_sheet.write(excel_row, ci + 1, 'N/A', prop_cell_fmt)
        
        # Column widths for the hidden sheet
        prop_data_sheet.set_column(0, 0, 30)
        for ci in range(len(compositions)):
            prop_data_sheet.set_column(ci + 1, ci + 1, 16)
        
        # Now create the Property Charts sheet
        prop_charts_sheet = workbook.add_worksheet('Property Charts')
        prop_charts_sheet.hide_gridlines(2)
        
        prop_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#3B82F6', 'font_name': 'Segoe UI'
        })
        prop_charts_sheet.write('B2', 'Physical, Mechanical & Optical Properties — Composition Charts', prop_title_fmt)
        
        # Scientific color palette matching matplotlib
        prop_palette = [
            '#3B82F6', '#F97316', '#2ca02c', '#d62728',
            '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
            '#bcbd22', '#17becf'
        ]
        
        prop_chart_row = 4  # Starting row in Property Charts sheet
        
        for p, cat in all_calculated_props:
            # Skip string-only properties — cannot be charted
            if p in STRING_ONLY_PROPS:
                continue
            excel_row = prop_row_map[p]
            
            # Determine unit label
            unit_str = ''
            for comp in compositions:
                col_full_name = f"{p} [{comp}]"
                try:
                    col_idx = headers.index(col_full_name)
                    unit_str = units[col_idx]
                    break
                except (ValueError, IndexError):
                    pass
            
            # Build a line chart with markers
            chart = workbook.add_chart({'type': 'line'})
            chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
            chart.set_plotarea({
                'border': {'color': '#cccccc', 'width': 1},
                'fill': {'color': '#ffffff'}
            })
            
            # Add one series per composition (each as a single point)
            # XlsxWriter line charts work best with all data in one series with category labels
            # We add a single series across all compositions
            start_col_letter = xlsxwriter.utility.xl_col_to_name(1)
            end_col_letter = xlsxwriter.utility.xl_col_to_name(len(compositions))
            
            chart.add_series({
                'name': p,
                'categories': f"=PropData!$B$1:${end_col_letter}$1",
                'values': f"=PropData!$B${excel_row + 1}:${end_col_letter}${excel_row + 1}",
                'line': {'color': prop_palette[0], 'width': 2.0},
                'marker': {
                    'type': 'circle',
                    'size': 7,
                    'fill': {'color': prop_palette[0]},
                    'border': {'color': '#ffffff', 'width': 1.5}
                },
                'data_labels': {
                    'value': True,
                    'num_format': '0.####',
                    'font': {'name': 'Segoe UI', 'size': 8, 'color': '#374151'},
                    'position': 'above'
                }
            })
            
            display_name = DISPLAY_NAMES.get(p, p)
            y_label = f"{display_name}{(' (' + unit_str + ')') if unit_str else ''}"
            
            chart.set_title({
                'name': f"{display_name}  [{cat}]",
                'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
            })
            chart.set_x_axis({
                'name': 'Composition Code',
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'bold': True, 'color': '#1f2937'},
                'major_gridlines': {'visible': False},
                'line': {'color': '#cccccc', 'width': 1},
            })
            valid_vals = []
            for comp in compositions:
                val = comp_props.get(comp, {}).get(p, (None,))[0]
                if val is not None and not pd.isna(val) and val != 'N/A':
                    try: valid_vals.append(float(val))
                    except: pass
            
            y_axis_config = {
                'name': y_label,
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'},
                'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}},
                'line': {'color': '#cccccc', 'width': 1},
            }
            y_axis_config.update(get_axis_bounds(valid_vals))
            chart.set_y_axis(y_axis_config)
            chart.set_legend({'none': True})
            chart.set_size({'width': 850, 'height': 420})
            
            prop_charts_sheet.insert_chart(f'B{prop_chart_row}', chart)
            prop_chart_row += 23
            
        # ── CLUBBED ELASTIC MODULI CHARTS ──
        elastic_moduli = ["Youngs Modulus", "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus"]
        present_moduli = [p for p, cat in all_calculated_props if p in elastic_moduli]
        
        if present_moduli:
            end_col_letter = xlsxwriter.utility.xl_col_to_name(len(compositions))
            
            # 1. Clubbed Line Chart
            clubbed_line = workbook.add_chart({'type': 'line'})
            clubbed_line.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
            clubbed_line.set_plotarea({
                'border': {'color': '#cccccc', 'width': 1},
                'fill': {'color': '#ffffff'}
            })
            
            for idx, p in enumerate(present_moduli):
                excel_row = prop_row_map[p]
                clubbed_line.add_series({
                    'name': p,
                    'categories': f"=PropData!$B$1:${end_col_letter}$1",
                    'values': f"=PropData!$B${excel_row + 1}:${end_col_letter}${excel_row + 1}",
                    'line': {'color': prop_palette[idx % len(prop_palette)], 'width': 2.0},
                    'marker': {
                        'type': 'circle',
                        'size': 7,
                        'fill': {'color': prop_palette[idx % len(prop_palette)]},
                        'border': {'color': '#ffffff', 'width': 1.5}
                    },
                    'data_labels': {
                        'value': True,
                        'num_format': '0.####',
                        'font': {'name': 'Segoe UI', 'size': 8, 'color': '#374151'},
                        'position': 'above'
                    }
                })
            
            clubbed_line.set_title({
                'name': 'Elastic Moduli Comparison (Line)',
                'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
            })
            clubbed_line.set_x_axis({
                'name': 'Composition Code',
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'bold': True, 'color': '#1f2937'},
                'major_gridlines': {'visible': False},
                'line': {'color': '#cccccc', 'width': 1},
            })
            # Get all values for scaling
            all_moduli_vals = []
            for p in present_moduli:
                for comp in compositions:
                    val = comp_props.get(comp, {}).get(p, (None,))[0]
                    if val is not None and not pd.isna(val) and val != 'N/A':
                        try: all_moduli_vals.append(float(val))
                        except: pass
                        
            y_axis_config = {
                'name': 'Elastic Moduli (GPa)',
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'},
                'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}},
                'line': {'color': '#cccccc', 'width': 1},
            }
            y_axis_config.update(get_axis_bounds(all_moduli_vals))
            clubbed_line.set_y_axis(y_axis_config)
            clubbed_line.set_legend({
                'position': 'bottom',
                'font': {'name': 'Segoe UI', 'size': 10, 'color': '#4b5563'}
            })
            clubbed_line.set_size({'width': 850, 'height': 420})
            prop_charts_sheet.insert_chart(f'B{prop_chart_row}', clubbed_line)
            prop_chart_row += 23
            
            # 2. Clubbed Bar Chart
            clubbed_bar = workbook.add_chart({'type': 'column'})
            clubbed_bar.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
            clubbed_bar.set_plotarea({
                'border': {'color': '#cccccc', 'width': 1},
                'fill': {'color': '#ffffff'}
            })
            
            for idx, p in enumerate(present_moduli):
                excel_row = prop_row_map[p]
                clubbed_bar.add_series({
                    'name': p,
                    'categories': f"=PropData!$B$1:${end_col_letter}$1",
                    'values': f"=PropData!$B${excel_row + 1}:${end_col_letter}${excel_row + 1}",
                    'fill': {'color': prop_palette[idx % len(prop_palette)]},
                    'data_labels': {
                        'value': True,
                        'num_format': '0.####',
                        'font': {'name': 'Segoe UI', 'size': 8, 'color': '#374151'},
                        'position': 'outside_end'
                    }
                })
                
            clubbed_bar.set_title({
                'name': 'Elastic Moduli Comparison (Bar)',
                'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
            })
            clubbed_bar.set_x_axis({
                'name': 'Composition Code',
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'bold': True, 'color': '#1f2937'},
                'major_gridlines': {'visible': False},
                'line': {'color': '#cccccc', 'width': 1},
            })
            clubbed_bar.set_y_axis({
                'name': 'Elastic Moduli (GPa)',
                'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                'num_font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'},
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}
                },
                'line': {'color': '#cccccc', 'width': 1},
            })
            clubbed_bar.set_legend({
                'position': 'bottom',
                'font': {'name': 'Segoe UI', 'size': 10, 'color': '#4b5563'}
            })
            clubbed_bar.set_size({'width': 850, 'height': 420})
            prop_charts_sheet.insert_chart(f'B{prop_chart_row}', clubbed_bar)
            prop_chart_row += 23
            
        # ── DOUBLE AXIS CHARTS ──
        double_axis_pairs = [
            ("Reflection Loss", "Transmission Coefficient", "Optical"),
            ("Density", "Molar Volume", "Physical"),
            ("Oxygen Packing Density", "Oxygen Molar Volume", "Physical")
        ]
        
        for prop1, prop2, cat in double_axis_pairs:
            if prop1 in prop_row_map and prop2 in prop_row_map:
                row1 = prop_row_map[prop1]
                row2 = prop_row_map[prop2]
                
                chart = workbook.add_chart({'type': 'line'})
                chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
                chart.set_plotarea({
                    'border': {'color': '#cccccc', 'width': 1},
                    'fill': {'color': '#ffffff'}
                })
                
                chart.add_series({
                    'name': prop1,
                    'categories': f"=PropData!$B$1:${end_col_letter}$1",
                    'values': f"=PropData!$B${row1 + 1}:${end_col_letter}${row1 + 1}",
                    'line': {'color': prop_palette[0], 'width': 2.0},
                    'marker': {
                        'type': 'circle', 'size': 7,
                        'fill': {'color': prop_palette[0]},
                        'border': {'color': '#ffffff', 'width': 1.5}
                    },
                    'data_labels': {
                        'value': True, 'num_format': '0.####',
                        'font': {'name': 'Segoe UI', 'size': 8, 'color': prop_palette[0]},
                        'position': 'above'
                    }
                })
                
                chart.add_series({
                    'name': prop2,
                    'categories': f"=PropData!$B$1:${end_col_letter}$1",
                    'values': f"=PropData!$B${row2 + 1}:${end_col_letter}${row2 + 1}",
                    'line': {'color': prop_palette[1], 'width': 2.0, 'dash_type': 'dash'},
                    'marker': {
                        'type': 'square', 'size': 7,
                        'fill': {'color': prop_palette[1]},
                        'border': {'color': '#ffffff', 'width': 1.5}
                    },
                    'y2_axis': True,
                    'data_labels': {
                        'value': True, 'num_format': '0.####',
                        'font': {'name': 'Segoe UI', 'size': 8, 'color': prop_palette[1]},
                        'position': 'below'
                    }
                })
                
                chart.set_title({
                    'name': f"{prop1} & {prop2} [{cat}]",
                    'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
                })
                
                chart.set_x_axis({
                    'name': 'Composition Code',
                    'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                    'num_font': {'name': 'Segoe UI', 'size': 9, 'bold': True, 'color': '#1f2937'},
                    'major_gridlines': {'visible': False},
                    'line': {'color': '#cccccc', 'width': 1},
                })
                
                def get_unit(p_name):
                    for comp in compositions:
                        col_full_name = f"{p_name} [{comp}]"
                        try: return units[headers.index(col_full_name)]
                        except: pass
                    return ""
                
                u1 = get_unit(prop1)
                u2 = get_unit(prop2)
                y1_label = f"{prop1}{(' (' + u1 + ')') if u1 else ''}"
                y2_label = f"{prop2}{(' (' + u2 + ')') if u2 else ''}"
                
                valid_y1 = []
                for comp in compositions:
                    val = comp_props.get(comp, {}).get(prop1, (None,))[0]
                    if val is not None and not pd.isna(val) and val != 'N/A':
                        try: valid_y1.append(float(val))
                        except: pass

                y1_config = {
                    'name': y1_label,
                    'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': prop_palette[0]},
                    'num_font': {'name': 'Segoe UI', 'size': 9, 'color': prop_palette[0]},
                    'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75, 'dash_type': 'solid'}},
                    'line': {'color': '#cccccc', 'width': 1},
                }
                y1_config.update(get_axis_bounds(valid_y1))
                chart.set_y_axis(y1_config)
                
                valid_y2 = []
                for comp in compositions:
                    val = comp_props.get(comp, {}).get(prop2, (None,))[0]
                    if val is not None and not pd.isna(val) and val != 'N/A':
                        try: valid_y2.append(float(val))
                        except: pass

                y2_config = {
                    'name': y2_label,
                    'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': prop_palette[1]},
                    'num_font': {'name': 'Segoe UI', 'size': 9, 'color': prop_palette[1]},
                    'line': {'color': '#cccccc', 'width': 1},
                }
                y2_config.update(get_axis_bounds(valid_y2))
                chart.set_y2_axis(y2_config)
                
                chart.set_legend({
                    'position': 'bottom',
                    'font': {'name': 'Segoe UI', 'size': 10, 'color': '#4b5563'}
                })
                chart.set_size({'width': 850, 'height': 420})
                
                prop_charts_sheet.insert_chart(f'B{prop_chart_row}', chart)
                prop_chart_row += 23

    # ── MULTI-THICKNESS GAMMA SHIELDING CHARTS (RPE, Leq, TF, EBF, EABF) ──
    # Column pattern: "RPE (t=1) [comp]", "Leq (t=2) [comp]", "TF (t=1) [comp]",
    #                 "EBF (x=1) [comp]", "EABF (x=2) [comp]"
    # Group by base metric (RPE, Leq, etc.) and composition, multi-lines per thickness/depth.
    thickness_metric_groups = {}   # base_metric -> {comp -> [(thickness_label, col_idx, sheet_name, num_rows)]}

    def collect_thickness_cols(dataframe, sheet_name):
        thick_pattern = re.compile(r'^(RPE|Leq|Transmission Factor|TF)\s*\(t=([^)]+)\)(?:\s*\[(.+)\])?$')
        depth_pattern  = re.compile(r'^(EBF|EABF)\s*\(x=([^)]+)\)(?:\s*\[(.+)\])?$')
        for col_idx, col in enumerate(dataframe.columns):
            for pat in (thick_pattern, depth_pattern):
                m = pat.match(col)
                if m:
                    metric    = m.group(1)   # e.g. "RPE"
                    param_val = m.group(2)   # e.g. "1"
                    comp_name = m.group(3) if m.group(3) else "Default"   # e.g. "0.1B2O3+0.9PbO"
                    thickness_metric_groups.setdefault(metric, {}).setdefault(comp_name, []).append(
                        (param_val, col_idx, sheet_name, len(dataframe))
                    )
                    break

    collect_thickness_cols(df_gamma, 'Gamma Data')
    if has_25:
        collect_thickness_cols(df_25, '25 Energies Data')

    if thickness_metric_groups:
        thick_chart_row = row_offset + 4  # start below regular charts

        for metric, comp_dict in thickness_metric_groups.items():
            for comp_name, entries in comp_dict.items():
                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
                chart.set_plotarea({'border': {'color': '#cccccc', 'width': 1}, 'fill': {'color': '#ffffff'}})

                for idx, (param_val, col_idx, sheet_name, num_rows) in enumerate(entries):
                    col_letter    = xlsxwriter.utility.xl_col_to_name(col_idx)
                    energy_letter = xlsxwriter.utility.xl_col_to_name(0)
                    sheet_ref     = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
                    series_color  = palette[idx % len(palette)]
                    x_label = 't' if metric in ('RPE', 'Leq', 'Transmission Factor', 'TF') else 'x'
                    display_metric = "Transmission Factor" if metric == "TF" else metric
                    series_label = f"{display_metric} ({x_label}={param_val} cm)"

                    chart.add_series({
                        'name': series_label,
                        'categories': f"={sheet_ref}!${energy_letter}$3:${energy_letter}${num_rows + 2}",
                        'values':     f"={sheet_ref}!${col_letter}$3:${col_letter}${num_rows + 2}",
                        'line':   {'color': series_color, 'width': 1.5},
                        'marker': {'type': 'circle', 'size': 4.5,
                                   'fill': {'color': series_color}, 'border': {'color': series_color}}
                    })

                x_label_full = 't' if metric in ('RPE', 'Leq', 'Transmission Factor', 'TF') else 'x (MFP)'
                display_metric = "Transmission Factor" if metric == "TF" else metric
                chart.set_title({
                    'name': f'{display_metric} — {comp_name} (multiple thicknesses)',
                    'name_font': {'name': 'Segoe UI', 'size': 13, 'bold': True, 'color': '#1f2937'}
                })
                chart.set_x_axis({
                    'name': 'Energy (MeV)',
                    'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                    'num_font':  {'name': 'Segoe UI', 'size': 9,  'color': '#4b5563'},
                    'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75}},
                    'log_base': 10,
                    'line': {'color': '#cccccc', 'width': 1},
                    'crossing': 'min',
                })
                y_log = metric in ('EBF', 'EABF')
                y_axis_cfg = {
                    'name': display_metric,
                    'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True, 'color': '#374151'},
                    'num_font':  {'name': 'Segoe UI', 'size': 9,  'color': '#4b5563'},
                    'major_gridlines': {'visible': True, 'line': {'color': '#e5e7eb', 'width': 0.75}},
                    'line': {'color': '#cccccc', 'width': 1},
                }
                if y_log:
                    y_axis_cfg['log_base'] = 10
                chart.set_y_axis(y_axis_cfg)
                chart.set_legend({'position': 'bottom', 'font': {'name': 'Segoe UI', 'size': 9, 'color': '#4b5563'}})
                chart.set_size({'width': 850, 'height': 500})

                charts_sheet.insert_chart(f'B{thick_chart_row}', chart)
                thick_chart_row += 27

    workbook.close()

    print(f"Workbook successfully saved to {xlsx_path}")

if __name__ == "__main__":
    main()
