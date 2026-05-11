import pandas as pd
import openpyxl

df = pd.read_excel('test_invoice.xlsx')
# df = df.iloc[0:7, 0:6]
print(df.head(10))

import pandas as pd
import numpy as np




def detect_table_bbox(df, min_row_density=0.3):

    # Convert everything to string safely
    temp = df.fillna("").astype(str)

    # Occupancy mask
    mask = temp.apply(lambda col: col.str.strip() != "")

    # Row density
    row_density = mask.sum(axis=1) / mask.shape[1]
    print("Row Density:", row_density)
    print(type(row_density))

    dense_rows = row_density >= min_row_density

    if not dense_rows.any():
        return None

    # Find contiguous dense row blocks
    blocks = []

    start = None

    for i, is_dense in enumerate(dense_rows):
        if is_dense and start==None:
            start = i
        
        elif not is_dense and start!=None:
            blocks.append((start, i-1))
            start=None

    if start!=None:
        blocks.append((start, len(blocks)-1))

    
    # Largest dense block
    start_row, end_row = max(blocks, key=lambda x: x[1] - x[0])

    # Detect active columns inside block
    submask = mask.iloc[start_row:end_row + 1]

    col_density = submask.sum(axis=0) / submask.shape[0]

    active_cols = np.where(col_density > 0)[0]

    if len(active_cols) == 0:
        return None

    start_col = active_cols[0]
    end_col = active_cols[-1]

    return (start_row, int(start_col), end_row, int(end_col))


bbox = detect_table_bbox(df)

print(bbox)