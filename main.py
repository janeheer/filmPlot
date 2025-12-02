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
journalCount = journalCount[~journalCount.index.str.contains('-')]

# Remove entries that aren't journal names
journalCount = journalCount[journalCount.index.str.len() > 5]

# Skip bad data!
ns_total = fundingCount.sum(skipna=True)
total = len(fundingCount)

# Pie of Funding Source
labels = ['specified', 'not specified']
plt.pie([ns_total, total - ns_total], labels=labels, autopct='%1.1f%%', startangle=90)
plt.show()

fig = plt.gcf()

# Save Plot
plt.savefig("FundingSourcePlot.png", dpi=300, bbox_inches='tight')
plt.close()

# Bar of Journal Names
fig, ax = plt.subplots()
ax.barh(journalCount.index, journalCount.values, align='edge', height=.8)
ax.set_xlabel('Number of Articles')
ax.set_title('Number of Articles by Journal Name')
ax.invert_yaxis()  # Highest values on top
ax.set_yticklabels(journalCount.index, fontsize=5)  

# Save Plot
fig = plt.gcf()
plt.savefig("JournalNamePlot.png", dpi=300, bbox_inches='tight')
plt.close()

# Bar Chart of Publication Years
publication_years = df['Publication Year'].dropna()
publication_years = publication_years[publication_years.apply(lambda x: str(x).isdigit())]
publication_years = publication_years.astype(int)
publication_years = publication_years.value_counts().sort_index()
fig, ax = plt.subplots()
ax.bar(publication_years.index.astype(str), publication_years.values, align='edge', width=0.8)

# Space out x-axis labels
ax.set_xticks(np.arange(len(publication_years.index)))
ax.set_xticklabels(publication_years.index.astype(str), rotation=45, ha='right')
ax.set_xlabel('Publication Year')
ax.set_ylabel('Number of Articles')
ax.set_title('Number of Articles by Publication Year')

# Save Plot
fig = plt.gcf()
plt.savefig("PublicationYearPlot.png", dpi=300, bbox_inches='tight')
plt.close() 

# Pull Data Deposited
data_deposited = df['Data Deposited? (Use Federer, et al 2018 categories)'].str.get_dummies(sep=';').sum()
labels = data_deposited.index
sizes = data_deposited.values

# Create bar chart for Data Deposited
fig, ax = plt.subplots()
ax.barh(labels, sizes, align='edge', height=.8)
ax.set_xlabel('Number of Articles')
ax.set_title('Number of Articles by Data Deposited Categories')
ax.invert_yaxis()  # Highest values on top
ax.set_yticklabels(labels, fontsize=5)

# Save Plot
fig = plt.gcf()
plt.savefig("DataDepositedPlot.png", dpi=300, bbox_inches='tight')
plt.close()