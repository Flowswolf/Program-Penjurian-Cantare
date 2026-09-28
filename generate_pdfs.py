import pandas as pd
import numpy as np
import os
from jinja2 import Environment, FileSystemLoader
import warnings

# Suppress pandas warning about openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Import xhtml2pdf for PDF generation
from xhtml2pdf import pisa

def get_medal(score):
    if pd.isna(score):
        return "NO"
    if 80 <= score <= 100:
        return "GOLD"
    elif 70 <= score < 80:
        return "SILVER"
    elif 60 <= score < 70:
        return "BRONZE"
    else:
        return "DIPLOMA"
        
def get_medal_color(medal):
    if medal == "GOLD":
        return "#FFD700"
    elif medal == "SILVER":
        return "#C0C0C0"
    elif medal == "BRONZE":
        return "#CD7F32"
    else:
        return "#E0E0E0"

def get_medal_text_color(medal):
    if medal == "SILVER":
        return "#333333"
    elif medal == "GOLD":
        return "#8B6508"
    elif medal == "BRONZE":
        return "#4A3C31"
    else:
        return "#333333"

def convert_html_to_pdf(source_html, output_filename):
    """
    Utility function to convert HTML string to PDF using xhtml2pdf
    """
    # Open file for writing
    with open(output_filename, "w+b") as result_file:
        # Convert HTML to PDF
        pisa_status = pisa.CreatePDF(
            source_html,
            dest=result_file
        )
    return not pisa_status.err

def process_excel(file_path):
    # Set up Jinja2 environment
    template_dir = os.path.join(os.path.dirname(__file__), 'template')
    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template('score_sheet.html')
    
    # Output directory
    output_dir = os.path.join(os.path.dirname(__file__), 'output_pdf')
    os.makedirs(output_dir, exist_ok=True)
    
    # Read all sheets from Excel
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    
    for sheet_name in xls.sheet_names:
        print(f"\nProcessing Category: {sheet_name}")
        
        try:
            # Read with 4 levels of headers (Title, Judge, Song, Metric)
            df = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1, 2, 3])
            
            # Find the column that contains 'Choir Name' in any of its levels
            choir_cols = [c for c in df.columns if "Choir Name" in c]
            
            if not choir_cols:
                # Try 3 levels (fallback for sample data)
                df = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1, 2])
                choir_cols = [c for c in df.columns if "Choir Name" in c]
                
            if not choir_cols:
                print(f"  -> Skipping sheet '{sheet_name}': 'Choir Name' column not found.")
                continue
                
        except Exception as e:
            print(f"  -> Skipping sheet '{sheet_name}': Could not parse headers ({str(e)})")
            continue
            
        # Determine Judges
        # If it's the real file (4 levels), judges are at level 1. If sample (3 levels), level 0.
        judge_level = 1 if len(df.columns[0]) == 4 else 0
        
        judges = []
        for col in df.columns.get_level_values(judge_level).unique():
            if pd.notna(col) and "Unnamed" not in str(col) and col not in ["Choir Name", "Penalty"]:
                judges.append(col)
                
        if not judges:
            print(f"  -> Skipping sheet '{sheet_name}': No judges found.")
            continue
            
        print(f"  Found Judges: {judges}")
        
        # Process each row (Choir)
        for index, row in df.iterrows():
            # Extract Choir Name (safely handling multiindex structure for Choir Name)
            choir_col = choir_cols[0]
            choir_name = row[choir_col]
            
            if pd.isna(choir_name) or str(choir_name).strip() == "":
                continue # Skip empty rows
                
            print(f"  Processing Choir: {choir_name}")
            
            list_juri_data = []
            total_scores_all_judges = []
            
            for judge in judges:
                # Extract judge's data
                try:
                    judge_data = row.xs(judge, level=judge_level)
                except KeyError:
                    continue # Judge not found in this row somehow
                
                # Calculate simple averages for the template display
                # I: Intonation
                i_vals = judge_data.xs('I', level=1, drop_level=False) if 'I' in judge_data.index.get_level_values(1) else pd.Series()
                # If there are multiple songs, it might be xs('I', level=1)
                
                # More robust way: find all I, F, VQ, AI
                i_vals = []
                f_vals = []
                vq_val = None
                ai_val = None
                all_valid_scores = []
                
                for idx_tuple, val in judge_data.items():
                    # If it's a multiindex, idx_tuple is a tuple. Otherwise just a scalar.
                    if isinstance(idx_tuple, tuple):
                        metric = str(idx_tuple[-1]).strip()
                    else:
                        metric = str(idx_tuple).strip()
                        
                    # Skip non-numeric values
                    if pd.notna(val) and isinstance(val, (int, float)):
                        all_valid_scores.append(val)
                        if metric == 'I': i_vals.append(val)
                        elif metric == 'F': f_vals.append(val)
                        elif metric == 'VQ': vq_val = val
                        elif metric == 'AI': ai_val = val
                
                # Criteria Averages
                avg_i = sum(i_vals) / len(i_vals) if i_vals else None
                avg_f = sum(f_vals) / len(f_vals) if f_vals else None
                
                # Calculate Judge's total score as average of criteria averages
                valid_criteria = []
                if avg_i is not None: valid_criteria.append(avg_i)
                if avg_f is not None: valid_criteria.append(avg_f)
                if vq_val is not None and pd.notna(vq_val): valid_criteria.append(vq_val)
                if ai_val is not None and pd.notna(ai_val): valid_criteria.append(ai_val)
                
                judge_total = sum(valid_criteria) / len(valid_criteria) if valid_criteria else 0
                total_scores_all_judges.append(judge_total)

                # Format to 2 decimal places if needed, or 0 if whole number
                fmt = lambda x: f"{x:.2f}".rstrip('0').rstrip('.') if x is not None and pd.notna(x) else ""
                
                list_juri_data.append({
                    "nama": judge,
                    "intonation": fmt(avg_i),
                    "fidelity": fmt(avg_f),
                    "vq": fmt(vq_val),
                    "artistic": fmt(ai_val),
                    "total": f"{judge_total:.2f}"
                })
                
            # Final calculations
            total_average = sum(total_scores_all_judges) / len(total_scores_all_judges) if total_scores_all_judges else 0
            
            # Extract penalty
            penalty_col = [c for c in df.columns if "Penalty" in c]
            penalty = 0
            if penalty_col and pd.notna(row[penalty_col[0]]):
                penalty = float(row[penalty_col[0]])
                
            final_score = total_average - penalty
            medal = get_medal(final_score)
            
            # Prepare data for Jinja
            template_data = {
                "nama_choir": choir_name,
                "kategori": sheet_name,
                "list_juri": list_juri_data,
                "total_average": f"{total_average:.2f}",
                "penalty": f"{penalty:.2f}" if penalty > 0 else "0",
                "final_score": f"{final_score:.2f}",
                "medal": medal,
                "medal_color": get_medal_color(medal),
                "medal_text_color": get_medal_text_color(medal)
            }
            
            # Render HTML
            html_content = template.render(**template_data)
            
            # Save to PDF
            safe_choir_name = str(choir_name).replace("/", "_").replace("\\", "_").replace(" ", "_")
            safe_category = str(sheet_name).split('(')[0].strip().replace(" ", "_")
            pdf_filename = f"{safe_category}_{safe_choir_name}.pdf"
            pdf_filepath = os.path.join(output_dir, pdf_filename)
            
            success = convert_html_to_pdf(html_content, pdf_filepath)
            if success:
                print(f"    -> Generated: {pdf_filename}")
            else:
                print(f"    -> Failed to generate: {pdf_filename}")

if __name__ == "__main__":
    real_file = "Penilaian Juri.xlsx"
    sample_file = "sample_scores.xlsx"
    
    if os.path.exists(real_file):
        print(f"Using real data file: {real_file}")
        process_excel(real_file)
        print("Done! Check the 'output_pdf' directory.")
    elif os.path.exists(sample_file):
        print(f"Using sample data file: {sample_file}")
        process_excel(sample_file)
        print("Done! Check the 'output_pdf' directory.")
    else:
        print(f"Neither {real_file} nor {sample_file} found. Please generate sample data first.")
