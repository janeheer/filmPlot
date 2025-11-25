import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
def spreadsheet_to_object(file_path, sheet_name=0):
    """
    Reads a spreadsheet and returns it as a Pandas DataFrame.
    
    :param file_path: Path to the spreadsheet (.xlsx, .xls, .csv)
    :param sheet_name: Sheet name or index (default: first sheet)
    :return: Pandas DataFrame
    """
    try:
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"File not found: {file_path}")
    except Exception as e:
        raise RuntimeError(f"Error reading spreadsheet: {e}")

# Extract
sheet = 'Coding_Template.xlsx'
df = spreadsheet_to_object(sheet)

# Count
fundingCount= df['Funding Source'].str.count('not specified')
journalCount= df['Journal Name'].value_counts()
print(journalCount)
ns_total = fundingCount.sum(skipna=True)
total = len(fundingCount)

# Pie of Funding Source
labels = ['specified', 'not specified']
plt.pie([ns_total, total - ns_total], labels=labels, autopct='%1.1f%%', startangle=90)
plt.show()

fig = plt.gcf()

plt.savefig("FundingSourcePlot.png", dpi=300, bbox_inches='tight')
plt.close()
# Bar of Journal Names
fig, ax = plt.subplots()
ax.barh(journalCount.index, journalCount.values, align='edge', height=.8)
ax.set_xlabel('Number of Articles')
ax.set_title('Number of Articles by Journal Name')
ax.invert_yaxis()  # Highest values on top
ax.set_yticklabels(journalCount.index, fontsize=5)  


fig = plt.gcf()
plt.savefig("JournalNamePlot.png", dpi=300, bbox_inches='tight')
plt.close()