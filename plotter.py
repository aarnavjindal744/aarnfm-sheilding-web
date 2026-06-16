import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import zipfile
import io
import re

matplotlib.rcParams['font.family'] = 'DejaVu Sans'

def is_comp_available_in_mech_db(comp_str):
    if not comp_str or comp_str == "Default":
        return True
    parts = comp_str.split('+')
    mech_allowed = {
        "V2O3", "GeO2", "B2O3", "V2O5", "SiO2", "TeO2", "As2O5", "Sb2O3", "Sb2O5", "P2O5",
        "Al2O3", "ZrO2", "BeO", "TiO2", "Ga2O3", "ThO2", "Fe2O3", "ZnO", "PbO", "CdO",
        "Ta2O5", "Cr2O3", "MgO", "Nb2O5", "Sc2O3", "Li2O", "Er2O3", "MnO2", "Y2O3", "Ce2O3",
        "Tm2O3", "MoO3", "Nd2O3", "Gd2O3", "CoO", "Pr2O3", "WO3", "Yb2O3", "La2O3", "CaO",
        "SnO2", "Dy2O3", "Sm2O3", "CuO", "Eu2O3", "SrO", "Cu2O", "SnO", "BaO", "PbO2",
        "Na2O", "Bi2O3", "In2O3", "Ag2O", "K2O", "Rb2O", "Cs2O"
    }
    for part in parts:
        part = part.strip()
        if not part:
            continue
        match = re.match(r'^[\d\.]*\s*(.*)$', part)
        if match:
            formula = match.group(1).strip()
            if formula and formula not in mech_allowed:
                return False
    return True

def main():
    if len(sys.argv) < 2:
        print("Usage: python plotter.py <output_zip_path>")
        sys.exit(1)

    zip_path = sys.argv[1]

    raw_data = sys.stdin.read()
    if not raw_data.strip():
        print("Error: No data provided to stdin")
        sys.exit(1)

    try:
        df = pd.read_csv(io.StringIO(raw_data), header=0, skiprows=[1, 2])
    except Exception as e:
        print(f"Error parsing CSV: {e}")
        sys.exit(1)

    if df.empty or len(df.columns) < 2:
        print("Error: DataFrame is empty or only has one column")
        sys.exit(1)

    energy_col = df.columns[0]

    # All physical/mechanical/optical scalar properties
    PROP_CHARTS = {
        "Molar Mass", "Molar Volume", "Avg Boron Distance", "Oxygen Packing Density",
        "Oxygen Molar Volume", "Ion Concentration", "Inter-ionic Distance", "Polaron Radius",
        "Field Strength",
        "Packing Density", "Dissociation Energy", "Youngs Modulus",
        "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus", "Poissons Ratio",
        "Fractal Bond Connectivity", "Vickers Hardness",
        "Refractive Index", "Dielectric Constant", "Optical Dielectric Constant",
        "Reflection Loss", "Molar Refractivity", "Metallization Criterion", "Molar Polarizability",
        "Electronic Polarizability", "Transmission Coefficient", "Optical Electronegativity",
        "Linear Susceptibility", "Non-linear Susceptibility", "Non-linear Refractive Index", 
        "Density", "CN", "Bandgap"
    }

    PROP_UNITS = {
        "Molar Mass": "g/mol", "Molar Volume": "cm³/mol",
        "Avg Boron Distance": "Å", "Oxygen Packing Density": "mol/L",
        "Oxygen Molar Volume": "cm³/mol", "Ion Concentration": "×10²¹ ions/cm³",
        "Inter-ionic Distance": "Å", "Polaron Radius": "Å",
        "Field Strength": "×10¹⁵ cm⁻²",
        "Packing Density": "", "Dissociation Energy": "kJ/cm³",
        "Youngs Modulus": "GPa", "Bulk Modulus": "GPa",
        "Shear Modulus": "GPa", "Longitudinal Modulus": "GPa",
        "Poissons Ratio": "", "Fractal Bond Connectivity": "", "Vickers Hardness": "GPa",
        "Refractive Index": "", "Dielectric Constant": "", "Optical Dielectric Constant": "",
        "Reflection Loss": "", "Molar Refractivity": "cm³/mol",
        "Metallization Criterion": "", "Molar Polarizability": "10⁻²⁴ cm³",
        "Electronic Polarizability": "10⁻²⁴ cm³", "Transmission Coefficient": "",
        "Optical Electronegativity": "", "Linear Susceptibility": "",
        "Non-linear Susceptibility": "10⁻¹ esu", "Non-linear Refractive Index": "10⁻¹⁵ esu",
        "Density": "g/cm³", "CN": "", "Bandgap": "eV"
    }

    # Display name overrides
    DISPLAY_NAMES = {"Bandgap": "Bandgap Energy"}

    CATEGORY = {
        "Molar Mass": "Physical", "Molar Volume": "Physical",
        "Avg Boron Distance": "Physical", "Oxygen Packing Density": "Physical",
        "Oxygen Molar Volume": "Physical", "Ion Concentration": "Physical",
        "Inter-ionic Distance": "Physical", "Polaron Radius": "Physical",
        "Field Strength": "Physical",
        "Packing Density": "Mechanical", "Dissociation Energy": "Mechanical",
        "Youngs Modulus": "Mechanical", "Bulk Modulus": "Mechanical",
        "Shear Modulus": "Mechanical", "Longitudinal Modulus": "Mechanical",
        "Poissons Ratio": "Mechanical", "Fractal Bond Connectivity": "Mechanical",
        "Vickers Hardness": "Mechanical",
        "Refractive Index": "Optical", "Dielectric Constant": "Optical",
        "Optical Dielectric Constant": "Optical", "Reflection Loss": "Optical",
        "Molar Refractivity": "Optical", "Metallization Criterion": "Optical",
        "Molar Polarizability": "Optical", "Electronic Polarizability": "Optical",
        "Transmission Coefficient": "Optical", "Optical Electronegativity": "Optical",
        "Linear Susceptibility": "Optical", "Non-linear Susceptibility": "Optical",
        "Non-linear Refractive Index": "Optical", "Density": "Physical",
        "CN": "Physical", "Bandgap": "Optical"
    }

    # Scientific color palette (tab10-inspired, same as xlsx)
    PALETTE = [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
        '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
        '#bcbd22', '#17becf'
    ]

    energy_groups = {}   # base_param -> [(comp_name, col)]
    prop_groups = {}     # base_param -> [(comp_name, value)]

    mechanical_props = {
        "Packing Density", "Dissociation Energy", "Youngs Modulus",
        "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus", "Poissons Ratio",
        "Fractal Bond Connectivity", "Vickers Hardness"
    }

    has_missing_mech_data = False
    for col in df.columns[1:]:
        if col == "Isotope":
            continue
        match = re.match(r"(.+?)\s*\[(.+?)\]", col)
        if match:
            base_param = match.group(1).strip()
            comp_name = match.group(2).strip()
        else:
            base_param = col
            comp_name = "Default"
        
        if base_param in mechanical_props:
            if not is_comp_available_in_mech_db(comp_name):
                has_missing_mech_data = True

    # Check if this is physical-only mode (dummy energy = 0.0 only)
    energy_col_numeric = pd.to_numeric(df[energy_col], errors='coerce')
    is_physical_only = (len(df) == 1 and energy_col_numeric.iloc[0] == 0.0) or \
                       (energy_col_numeric.dropna().empty)

    for col in df.columns[1:]:
        if col == "Isotope":
            continue

        df[col] = pd.to_numeric(df[col], errors='coerce')
        if not pd.api.types.is_numeric_dtype(df[col]):
            continue

        match = re.match(r"(.+?)\s*\[(.+?)\]", col)
        if match:
            base_param = match.group(1).strip()
            comp_name = match.group(2).strip()
        else:
            base_param = col
            comp_name = "Default"

        if base_param == "Dopant":
            continue

        if base_param in PROP_CHARTS:
            series_nonan = df[col].dropna()
            val = series_nonan.iloc[0] if not series_nonan.empty else float('nan')
            if base_param not in prop_groups:
                prop_groups[base_param] = []
            prop_groups[base_param].append((comp_name, val))
        else:
            # Skip energy charts in physical-only mode
            if is_physical_only:
                continue
            if base_param not in energy_groups:
                energy_groups[base_param] = []
            energy_groups[base_param].append((comp_name, col))

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:

        # ── 1. Energy-based line charts ──
        for base_param, members in energy_groups.items():
            # Skip thickness-variant params — handled separately below
            if re.match(r'^(RPE|Leq|Transmission Factor|EBF|EABF)\s*\(', base_param):
                continue
            fig, ax = plt.subplots(figsize=(10, 7))
            fig.patch.set_facecolor('#f8f9fa')
            ax.set_facecolor('#ffffff')

            for i, (comp_name, col) in enumerate(members):
                label = comp_name if comp_name != "Default" else base_param
                color = PALETTE[i % len(PALETTE)]

                valid_count = df[col].notna().sum()
                if valid_count <= 25 and len(df) > 25 and valid_count > 0:
                    non_na_mask = df[col].notna()
                    x_vals = df.loc[non_na_mask, energy_col].values
                    y_vals = df.loc[non_na_mask, col].values
                    ax.plot(x_vals, y_vals, marker='o', linestyle='-',
                            markersize=4, linewidth=1.5, label=label, color=color)
                else:
                    ax.plot(df[energy_col], df[col], marker='o', linestyle='-',
                            markersize=4, linewidth=1.5, label=label, color=color)

            ax.set_xlabel(f"{energy_col} (MeV)", fontsize=12, color='#374151')
            ax.set_ylabel(base_param, fontsize=12, color='#374151')
            ax.set_title(f"{base_param} vs {energy_col}", fontsize=14, fontweight='bold', color='#1f2937')

            if len(members) > 1 or members[0][0] != "Default":
                ax.legend(loc="best", fontsize=10, framealpha=0.8)

            if any(x in base_param for x in ["MAC", "LAC", "MFP", "HVL", "TVL", "EBF", "EABF"]):
                ax.set_xscale('log')
                ax.set_yscale('log')
            else:
                ax.set_xscale('log')

            ax.grid(True, which="both", ls="--", alpha=0.5, color='#d1d5db')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.tight_layout()

            buf = io.BytesIO()
            plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
            plt.close()

            safe_name = "".join([c if c.isalnum() else "_" for c in base_param]) + ".jpeg"
            zf.writestr(safe_name, buf.getvalue())

        # ── 2. Composition line charts for physical/mechanical/optical properties ──
        # Styled to match the xlsx Property Charts sheet
        for base_param, members in prop_groups.items():
            if len(members) < 1:
                continue

            if base_param in mechanical_props and has_missing_mech_data:
                continue

            comp_names = [m[0] for m in members]
            values = [m[1] for m in members]
            n = len(comp_names)

            fig_w = max(7, n * 1.4 + 2)
            fig, ax = plt.subplots(figsize=(fig_w, 5.5))
            fig.patch.set_facecolor('#ffffff')
            ax.set_facecolor('#ffffff')

            # Connecting line (gray, like xlsx)
            ax.plot(range(n), values, linestyle='-', linewidth=2.0,
                    color='#1f77b4', alpha=0.75, zorder=2)

            # One marker per composition — colored by palette (like xlsx series markers)
            for i, (name, val) in enumerate(zip(comp_names, values)):
                ax.plot(i, val,
                        marker='o', markersize=10,
                        color=PALETTE[i % len(PALETTE)],
                        markeredgecolor='white', markeredgewidth=1.5,
                        zorder=3)

            # Data labels above each point (matching xlsx data_labels: value=True)
            for i, (name, val) in enumerate(zip(comp_names, values)):
                if pd.notna(val):
                    ax.annotate(f"{val:.4g}",
                                xy=(i, val), xytext=(0, 11),
                                textcoords='offset points',
                                ha='center', fontsize=9, fontweight='600',
                                color='#1f2937')

            display_name = DISPLAY_NAMES.get(base_param, base_param)
            unit = PROP_UNITS.get(base_param, "")
            cat = CATEGORY.get(base_param, "")
            ylabel = f"{display_name}{(' (' + unit + ')') if unit else ''}"

            ax.set_xticks(range(n))
            ax.set_xticklabels(comp_names, fontsize=11, fontweight='700', color='#1f2937')
            ax.set_ylabel(ylabel, fontsize=11, color='#374151')
            ax.set_title(f"{display_name}  [{cat}]\nComposition Comparison",
                         fontsize=13, fontweight='bold', color='#1f2937', pad=16)

            ax.margins(x=0.2)
            ax.grid(True, axis='y', ls='--', alpha=0.45, color='#e5e7eb')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_color('#cccccc')
            ax.spines['left'].set_color('#cccccc')
            ax.tick_params(axis='x', length=0)

            plt.tight_layout()

            buf = io.BytesIO()
            plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
            plt.close()

            safe_name = "PROP_" + "".join([c if c.isalnum() else "_" for c in base_param]) + ".jpeg"
            zf.writestr(safe_name, buf.getvalue())

        # ── 2.5 Clubbed Elastic Moduli comparison graphs (Line and Bar) ──
        elastic_moduli = ["Youngs Modulus", "Bulk Modulus", "Shear Modulus", "Longitudinal Modulus"]
        present_moduli = [p for p in elastic_moduli if p in prop_groups]
        
        if present_moduli and not has_missing_mech_data:
            all_comps = []
            for p in present_moduli:
                for comp_name, val in prop_groups[p]:
                    if comp_name not in all_comps:
                        all_comps.append(comp_name)
            
            n = len(all_comps)
            if n > 0:
                # 1. Clubbed Line Chart
                fig_w = max(8, n * 1.4 + 2)
                fig, ax = plt.subplots(figsize=(fig_w, 6))
                fig.patch.set_facecolor('#ffffff')
                ax.set_facecolor('#ffffff')
                
                for idx, p in enumerate(present_moduli):
                    val_dict = {comp: val for comp, val in prop_groups[p]}
                    values = [val_dict.get(c, float('nan')) for c in all_comps]
                    color = PALETTE[idx % len(PALETTE)]
                    
                    ax.plot(range(n), values, linestyle='-', linewidth=2.0,
                            color=color, alpha=0.85, label=p, zorder=2)
                    
                    for i, val in enumerate(values):
                        if pd.notna(val):
                            ax.plot(i, val, marker='o', markersize=8,
                                    color=color, markeredgecolor='white', markeredgewidth=1.2,
                                    zorder=3)
                            ax.annotate(f"{val:.4g}", xy=(i, val), xytext=(0, 9),
                                        textcoords='offset points', ha='center',
                                        fontsize=8, fontweight='600', color='#374151')
                
                ax.set_xticks(range(n))
                ax.set_xticklabels(all_comps, fontsize=11, fontweight='700', color='#1f2937')
                ax.set_ylabel("Elastic Moduli (GPa)", fontsize=11, color='#374151')
                ax.set_title("Elastic Moduli Comparison\nComposition Comparison (Line)",
                             fontsize=13, fontweight='bold', color='#1f2937', pad=16)
                ax.legend(loc="best", fontsize=10, framealpha=0.8)
                ax.margins(x=0.2)
                ax.grid(True, axis='y', ls='--', alpha=0.45, color='#e5e7eb')
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_color('#cccccc')
                ax.spines['left'].set_color('#cccccc')
                ax.tick_params(axis='x', length=0)
                plt.tight_layout()
                
                buf = io.BytesIO()
                plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
                plt.close()
                zf.writestr("PROP_Elastic_Moduli_Comparison_Line.jpeg", buf.getvalue())
                
                # 2. Clubbed Bar Chart
                fig, ax = plt.subplots(figsize=(fig_w, 6))
                fig.patch.set_facecolor('#ffffff')
                ax.set_facecolor('#ffffff')
                
                width = 0.8 / len(present_moduli)
                x_indices = list(range(n))
                
                for idx, p in enumerate(present_moduli):
                    val_dict = {comp: val for comp, val in prop_groups[p]}
                    values = [val_dict.get(c, float('nan')) for c in all_comps]
                    color = PALETTE[idx % len(PALETTE)]
                    
                    offset = (idx - (len(present_moduli) - 1) / 2) * width
                    bar_x = [x + offset for x in x_indices]
                    
                    rects = ax.bar(bar_x, [v if pd.notna(v) else 0.0 for v in values], width,
                                   label=p, color=color, edgecolor='none', zorder=2)
                    
                    for i, v in enumerate(values):
                        if pd.notna(v):
                            ax.annotate(f"{v:.4g}",
                                        xy=(bar_x[i], v),
                                        xytext=(0, 3),
                                        textcoords="offset points",
                                        ha='center', va='bottom', fontsize=8,
                                        fontweight='600', color='#374151')
                
                ax.set_xticks(x_indices)
                ax.set_xticklabels(all_comps, fontsize=11, fontweight='700', color='#1f2937')
                ax.set_ylabel("Elastic Moduli (GPa)", fontsize=11, color='#374151')
                ax.set_title("Elastic Moduli Comparison\nComposition Comparison (Bar)",
                             fontsize=13, fontweight='bold', color='#1f2937', pad=16)
                ax.legend(loc="best", fontsize=10, framealpha=0.8)
                ax.grid(True, axis='y', ls='--', alpha=0.45, color='#e5e7eb')
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_color('#cccccc')
                ax.spines['left'].set_color('#cccccc')
                ax.tick_params(axis='x', length=0)
                plt.tight_layout()
                
                buf = io.BytesIO()
                plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
                plt.close()
                zf.writestr("PROP_Elastic_Moduli_Comparison_Bar.jpeg", buf.getvalue())

        # ── 2.6 Double-Axis Comparison Graphs ──
        double_axis_pairs = [
            ("Reflection Loss", "Transmission Coefficient", "Optical"),
            ("Density", "Molar Volume", "Physical"),
            ("Oxygen Packing Density", "Oxygen Molar Volume", "Physical")
        ]
        
        for prop1, prop2, cat in double_axis_pairs:
            if prop1 in prop_groups and prop2 in prop_groups:
                all_comps = []
                for comp_name, _ in prop_groups[prop1]:
                    if comp_name not in all_comps: all_comps.append(comp_name)
                for comp_name, _ in prop_groups[prop2]:
                    if comp_name not in all_comps: all_comps.append(comp_name)
                
                n = len(all_comps)
                if n > 0:
                    fig_w = max(8, n * 1.4 + 3)
                    fig, ax1 = plt.subplots(figsize=(fig_w, 6))
                    fig.patch.set_facecolor('#ffffff')
                    ax1.set_facecolor('#ffffff')
                    
                    val_dict1 = {comp: val for comp, val in prop_groups[prop1]}
                    values1 = [val_dict1.get(c, float('nan')) for c in all_comps]
                    
                    val_dict2 = {comp: val for comp, val in prop_groups[prop2]}
                    values2 = [val_dict2.get(c, float('nan')) for c in all_comps]
                    
                    color1 = PALETTE[0]
                    color2 = PALETTE[1]
                    
                    ax1.plot(range(n), values1, linestyle='-', linewidth=2.0, color=color1, label=prop1, zorder=2)
                    for i, val in enumerate(values1):
                        if pd.notna(val):
                            ax1.plot(i, val, marker='o', markersize=8, color=color1, markeredgecolor='white', markeredgewidth=1.2, zorder=3)
                            ax1.annotate(f"{val:.4g}", xy=(i, val), xytext=(0, 9), textcoords='offset points', ha='center', fontsize=8, fontweight='600', color=color1)
                    
                    ax2 = ax1.twinx()
                    ax2.plot(range(n), values2, linestyle='--', linewidth=2.0, color=color2, label=prop2, zorder=2)
                    for i, val in enumerate(values2):
                        if pd.notna(val):
                            ax2.plot(i, val, marker='s', markersize=8, color=color2, markeredgecolor='white', markeredgewidth=1.2, zorder=3)
                            ax2.annotate(f"{val:.4g}", xy=(i, val), xytext=(0, -15), textcoords='offset points', ha='center', fontsize=8, fontweight='600', color=color2)
                    
                    ax1.set_xticks(range(n))
                    ax1.set_xticklabels(all_comps, fontsize=11, fontweight='700', color='#1f2937')
                    
                    unit1 = PROP_UNITS.get(prop1, "")
                    unit2 = PROP_UNITS.get(prop2, "")
                    y1_label = f"{prop1}{(' (' + unit1 + ')') if unit1 else ''}"
                    y2_label = f"{prop2}{(' (' + unit2 + ')') if unit2 else ''}"
                    
                    ax1.set_ylabel(y1_label, fontsize=11, color=color1)
                    ax2.set_ylabel(y2_label, fontsize=11, color=color2)
                    
                    plt.title(f"{prop1} & {prop2} [{cat}]\nDouble-Axis Comparison", fontsize=13, fontweight='bold', color='#1f2937', pad=16)
                    
                    lines_1, labels_1 = ax1.get_legend_handles_labels()
                    lines_2, labels_2 = ax2.get_legend_handles_labels()
                    ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc="best", fontsize=10, framealpha=0.8)
                    
                    ax1.grid(True, axis='y', ls='--', alpha=0.45, color='#e5e7eb')
                    ax1.spines['top'].set_visible(False)
                    ax2.spines['top'].set_visible(False)
                    
                    plt.tight_layout()
                    buf = io.BytesIO()
                    plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
                    plt.close()
                    
                    safe_name1 = "".join([c if c.isalnum() else "_" for c in prop1])
                    safe_name2 = "".join([c if c.isalnum() else "_" for c in prop2])
                    zf.writestr(f"PROP_DoubleAxis_{safe_name1}_{safe_name2}.jpeg", buf.getvalue())

        # ── 2.7 Multi-thickness gamma shielding charts ──
        # Detect columns like "RPE (t=1) [comp]", "EBF (x=5) [comp]"
        thick_pattern = re.compile(r'^(RPE|Leq|Transmission Factor|TF)\s*\(t=([^)]+)\)(?:\s*\[(.+)\])?$')
        depth_pattern  = re.compile(r'^(EBF|EABF)\s*\(x=([^)]+)\)(?:\s*\[(.+)\])?$')
        thick_groups = {}  # (metric, comp) -> [(param_val, col)]

        for col in df.columns[1:]:
            for pat in (thick_pattern, depth_pattern):
                m = pat.match(col)
                if m:
                    metric    = m.group(1)
                    param_val = m.group(2)
                    comp_name = m.group(3) if m.group(3) else "Default"
                    key = (metric, comp_name)
                    thick_groups.setdefault(key, []).append((param_val, col))
                    break

        for (metric, comp_name), entries in thick_groups.items():
            fig, ax = plt.subplots(figsize=(10, 7))
            fig.patch.set_facecolor('#f8f9fa')
            ax.set_facecolor('#ffffff')

            x_label = 't' if metric in ('RPE', 'Leq', 'Transmission Factor', 'TF') else 'x'
            unit    = '%' if metric == 'RPE' else ('cm' if metric == 'Leq' else '')
            display_metric = "Transmission Factor" if metric == "TF" else metric

            for i, (param_val, col) in enumerate(entries):
                color = PALETTE[i % len(PALETTE)]
                series = pd.to_numeric(df[col], errors='coerce')
                energy_numeric = pd.to_numeric(df[energy_col], errors='coerce')

                valid_mask = series.notna() & energy_numeric.notna()
                x_vals = energy_numeric[valid_mask].values
                y_vals = series[valid_mask].values

                if len(x_vals) > 0:
                    ax.plot(x_vals, y_vals, marker='o', linestyle='-',
                            markersize=4, linewidth=1.5,
                            label=f'{display_metric} ({x_label}={param_val} cm)',
                            color=color)

            ax.set_xscale('log')
            if metric in ('EBF', 'EABF'):
                ax.set_yscale('log')

            y_axis_label = f"{display_metric}{(' (' + unit + ')') if unit else ''}"
            ax.set_xlabel('Energy (MeV)', fontsize=12, color='#374151')
            ax.set_ylabel(y_axis_label, fontsize=12, color='#374151')
            ax.set_title(f'{display_metric} — {comp_name}\n(Multiple Thicknesses)',
                         fontsize=14, fontweight='bold', color='#1f2937')
            ax.legend(loc='best', fontsize=10, framealpha=0.8)
            ax.grid(True, which='both', ls='--', alpha=0.5, color='#d1d5db')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.tight_layout()

            buf = io.BytesIO()
            plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
            plt.close()

            safe_metric = ''.join([c if c.isalnum() else '_' for c in metric])
            safe_comp   = ''.join([c if c.isalnum() else '_' for c in comp_name])
            zf.writestr(f'THICK_{safe_metric}_{safe_comp}.jpeg', buf.getvalue())

        # ── 3. Overview gallery by category ──

        for cat_name in ["Physical", "Mechanical", "Optical"]:
            if cat_name == "Mechanical" and has_missing_mech_data:
                continue
            cat_props = [(p, v) for p, v in prop_groups.items() if CATEGORY.get(p) == cat_name]
            if not cat_props:
                continue

            n_props = len(cat_props)
            ncols = min(3, n_props)
            nrows = (n_props + ncols - 1) // ncols

            fig, axes = plt.subplots(nrows, ncols,
                                     figsize=(ncols * 5, nrows * 4 + 1),
                                     squeeze=False)
            fig.patch.set_facecolor('#ffffff')
            fig.suptitle(f"{cat_name} Properties — Composition Comparison",
                         fontsize=14, fontweight='bold', color='#1f2937', y=1.01)

            axes_flat = [axes[r][c] for r in range(nrows) for c in range(ncols)]

            for ax_idx, (base_param, members) in enumerate(cat_props):
                ax = axes_flat[ax_idx]
                ax.set_facecolor('#ffffff')

                names = [m[0] for m in members]
                vals = [m[1] for m in members]
                nn = len(names)

                ax.plot(range(nn), vals, linestyle='-', linewidth=1.8,
                        color='#1f77b4', alpha=0.65, zorder=2)
                for i, (nm, vl) in enumerate(zip(names, vals)):
                    ax.plot(i, vl, marker='o', markersize=8,
                            color=PALETTE[i % len(PALETTE)],
                            markeredgecolor='white', markeredgewidth=1.2, zorder=3)
                    if pd.notna(vl):
                        ax.annotate(f"{vl:.3g}", xy=(i, vl), xytext=(0, 9),
                                    textcoords='offset points',
                                    ha='center', fontsize=7.5, fontweight='600',
                                    color='#374151')

                unit = PROP_UNITS.get(base_param, "")
                ax.set_xticks(range(nn))
                ax.set_xticklabels(names, fontsize=8.5, fontweight='700')
                ax.set_title(base_param, fontsize=9.5, fontweight='bold', color='#374151')
                if unit:
                    ax.set_ylabel(unit, fontsize=8, color='#6b7280')
                ax.grid(True, axis='y', ls='--', alpha=0.4, color='#e5e7eb')
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_color('#cccccc')
                ax.spines['left'].set_color('#cccccc')
                ax.tick_params(axis='x', length=0)

            for ax_idx in range(n_props, len(axes_flat)):
                axes_flat[ax_idx].set_visible(False)

            plt.tight_layout()

            buf = io.BytesIO()
            plt.savefig(buf, format='jpeg', dpi=150, bbox_inches='tight')
            plt.close()

            zf.writestr(f"OVERVIEW_{cat_name}_Properties.jpeg", buf.getvalue())

        if has_missing_mech_data:
            zf.writestr("mechanical_properties_error.txt", "compound is not available in database(MECHANICAL)")

if __name__ == "__main__":
    main()
