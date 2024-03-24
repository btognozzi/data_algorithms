import pandas as pd
import numpy as np
import string

def clean_text(text):
    if isinstance(text, str):
        # Remove punctuation and newline characters for strings
        translator = str.maketrans('', '', string.punctuation)
        return text.translate(translator).strip()
    else:
        # Return the original value for non-strings (e.g., numbers)
        return text

def classify_result(row):
    if row['Sample ID'] == 'NA13591':
        if row['Assay Name'] == 'H63D':
            return '' if row['Call'] == 'MUTMUT' else 'NA13591 control has failed'
        elif row['Assay Name'] == 'C282Y':
            return 'NA13591 control has passed' if row['Call'] == 'WTWT' else 'NA13591 control has failed'
    elif row['Sample ID'] == 'NA14640':
        if row['Assay Name'] == 'H63D':
            return '' if row['Call'] == 'WTWT' else 'NA14640 control has failed'
        elif row['Assay Name'] == 'C282Y':
            return 'NA14640 control has passed' if row['Call'] == 'MUTMUT' else 'NA14640 control has failed'
    elif row['Sample ID'] == 'NA14691':
        if row['Assay Name'] == 'H63D':
            return '' if row['Call'] == 'WTMUT' else 'NA14691 control has failed'
        elif row['Assay Name'] == 'C282Y':
            return 'NA14691 control has passed' if row['Call'] == 'WTMUT' else 'NA14691 control has failed'
    elif row['Sample ID'] == 'NTC':
        if row['Assay Name'] == 'H63D':
            return '' if row['Call'] == '' else 'NTC control has failed'
        elif row['Assay Name'] == 'C282Y':
            return 'NTC control has passed' if row['Call'] == '' else 'NTC control has failed'
    elif row['Assay Name'] == 'C282Y':
        if row['Call'] == 'WTWT':
            return f"{row['Sample ID']} is negative for C282Y"
        elif row['Call'] == 'WTMUT':
            return f"{row['Sample ID']} is het for C282Y"
        elif row['Call'] == 'MUTMUT':
            return f"{row['Sample ID']} is homo for C282Y"
        elif row['Call'] == 'NOAMP':
            return f"{row['Sample ID']} is NO AMP for C282Y"
        else:
            return f"{row['Sample ID']} is inconclusive for C282Y"
    elif row['Assay Name'] == 'H63D':
        if row['Call'] == 'WTWT':
            return 'and negative for H63D'
        elif row['Call'] == 'WTMUT':
            return 'and het for H63D'
        elif row['Call'] == 'MUTMUT':
            return 'and homo for H63D'
        elif row['Call'] == 'NOAMP':
            return "and is NO AMP for H63D"
        else:
            return 'and inconclusive for H63D'
    else:
        return 'Assay is Undefined'

txt_file = input('Input Text File Path:')
df = pd.read_csv(txt_file, sep='\t', skiprows=17)
df.drop(df.index[-21:], axis = 0, inplace = True)
df = df.replace(np.nan, '', regex=True)
df.sort_values(by = ['Assay Name'], ascending = True, inplace = True)

# Clean the text in the dataframe
df = df.map(clean_text)

# Cleaned text in data frame, before final processing
df_pre = df.copy()
df_pre.drop(df.columns[6:], axis = 1, inplace = True)

# Apply the classification function to each row and create a new column
df['Classification'] = df.apply(classify_result, axis=1)

# Group by "Sample ID" and aggregate the classifications into a single line
result_df = df.groupby('Sample ID')['Classification'].apply(lambda x: ' '.join(x)).reset_index()

# Save the results to an Excel file
output = input('Input Output File Path:')
with pd.ExcelWriter(output) as writer:
    df.to_excel(writer, sheet_name = 'Raw Data', index = False)
    df_pre.to_excel(writer, sheet_name = 'Pre Processed', index = False)
    result_df.to_excel(writer, sheet_name = 'Results', index = False)

# Display the classifications for each Sample ID
for index, row in result_df.iterrows():
    print(f"{row['Classification']}")
