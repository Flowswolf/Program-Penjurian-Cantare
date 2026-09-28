import pandas as pd
import numpy as np

def create_sample_excel(filename="sample_scores.xlsx"):
    # Create an empty excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # --- SHEET 1: Early Children (A0) - 2 Songs ---
        # Create columns according to the screenshot
        
        # 5 Judges, 8 columns each (but Song 3 is blank for 2 song categories)
        # However, the structure is:
        # Col A: empty (or index)
        # Col B: Choir Name
        # Judge 1: C to J
        # Judge 2: K to R
        # Judge 3: S to Z
        # Judge 4: AA to AH
        # Judge 5: AI to AP
        # Penalty: AQ
        
        # To recreate the multi-level header:
        # Row 1: Category Name
        # Row 2: Judge Names (merged)
        # Row 3: Song 1, Song 2, Song 3 (merged)
        # Row 4: I, F, I, F, I, F, VQ, AI
        
        # We will build a simple DataFrame that mimics this structure.
        # Pandas MultiIndex is best for this.
        
        judges = [
            "Lim Ai Hooi", 
            "Lee Shiak Yao", 
            "Dimitris Ktistakis", 
            "Budi Utomo Prabowo", 
            "Andrea Angelini"
        ]
        
        columns = [("Choir Name", "", "")]
        
        for judge in judges:
            columns.extend([
                (judge, "Song 1", "I"),
                (judge, "Song 1", "F"),
                (judge, "Song 2", "I"),
                (judge, "Song 2", "F"),
                (judge, "Song 3", "I"),
                (judge, "Song 3", "F"),
                (judge, "", "VQ"),
                (judge, "", "AI"),
            ])
            
        columns.append(("Penalty", "", ""))
        
        multi_col = pd.MultiIndex.from_tuples(columns)
        
        # Create some dummy data for Early Children (A0)
        data_a0 = [
            ["KB TK Islam Al Azhar 35"] + 
            [78, 79, 78, 79, np.nan, np.nan, 78, 78] + # Judge 1
            [80, 80, 80, 80, np.nan, np.nan, 78, 78] + # Judge 2
            [82, 82, 81, 82, np.nan, np.nan, 84, 84] + # Judge 3
            [80, 75, 77, 78, np.nan, np.nan, 78, 80] + # Judge 4
            [75, 77, 72, 75, np.nan, np.nan, 72, 74] + # Judge 5
            [0], # Penalty
            
            ["Petra Early Children"] + 
            [80, 81, 80, 81, np.nan, np.nan, 80, 80] + 
            [82, 82, 82, 82, np.nan, np.nan, 80, 80] + 
            [84, 84, 83, 84, np.nan, np.nan, 86, 86] + 
            [82, 77, 79, 80, np.nan, np.nan, 80, 82] + 
            [77, 79, 74, 77, np.nan, np.nan, 74, 76] + 
            [3], # Penalty
        ]
        
        df_a0 = pd.DataFrame(data_a0, columns=multi_col)
        df_a0.to_excel(writer, sheet_name="Early Children (A0)")


        # --- SHEET 2: Mixed Youth (B1) - 3 Songs ---
        data_b1 = [
            ["Adi Nugroho Youth Choir"] + 
            [73, 73, 72, 71, 68, 68, 70, 71] + # Judge 1
            [65, 65, 65, 65, 63, 63, 70, 67] + # Judge 2
            [72, 65, 67, 68, 67, 67, 67, 71] + # Judge 3
            [63, 60, 62, 60, 68, 60, 60, 60] + # Judge 4
            [72, 72, 68, 70, 65, 65, 65, 68] + # Judge 5
            [0], # Penalty
            
            ["Petra Youth Choir"] + 
            [82, 82, 82, 82, 81, 81, 82, 82] + 
            [77, 77, 77, 75, 77, 77, 82, 77] + 
            [77, 80, 78, 81, 77, 79, 78, 79] + 
            [77, 72, 75, 70, 78, 77, 77, 76] + 
            [75, 76, 75, 76, 74, 76, 75, 77] + 
            [0], # Penalty
        ]
        
        df_b1 = pd.DataFrame(data_b1, columns=multi_col)
        df_b1.to_excel(writer, sheet_name="Mixed Youth (B1)")

    print(f"Sample data created successfully in {filename}")

if __name__ == "__main__":
    create_sample_excel()
