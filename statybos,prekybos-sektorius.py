import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

print("All libraries imported successfully!")



# Read file from project directory
df = pd.read_excel('APS_dalyviai_duomenys iš VMI.xlsx', header=0)
df = df.dropna(how='all')

def is_all_strings(row):
    return all(isinstance(item, str) for item in row)

# Filter out rows where all values are strings
df = df[~df.apply(is_all_strings, axis=1)]

# Reset the index after filtering
df.reset_index(drop=True, inplace=True)

# Split the DataFrame into four parts, each with 201 rows
q1 = df.iloc[:200]
q2 = df.iloc[200:400]
q3 = df.iloc[400:600]
q4 = df.iloc[601:800]

AllQuarters = pd.concat([q1, q2, q3, q4])

def filter_rows_by_starting_letter(df, starting_letter):

    if 'EVRK' not in df.columns:
        raise ValueError("DataFrame must contain an 'EVRK' column.")

    if not isinstance(starting_letter, str) or len(starting_letter) != 1:
        raise ValueError("Starting letter must be a single character string.")

    filtered_df = df[df['EVRK'].astype(str).str.startswith(starting_letter)]
    return filtered_df

Statybos =filter_rows_by_starting_letter(AllQuarters, "F")
Prekyba =filter_rows_by_starting_letter(AllQuarters, "G")

print(Statybos.head())

# Duomenu valymas:

def filter_data(df):
    # Step 1: Drop unwanted columns
    columns_to_exclude = ['ID_N', 'buvo_aps_dalyviu' , 'EVRK']
    filtered_df = df.drop(columns=columns_to_exclude, errors='ignore')  # 'errors=ignore' skips if columns don't exist

    # Step 2: Convert 'n/d' to NaN and ensure numeric data
    numeric_df = filtered_df.apply(pd.to_numeric, errors='coerce') # 'n/d' → NaN

    # Step 3: Filter out values > |1000| in specific columns
    for col in numeric_df.columns:
        if col.endswith('_POKYTIS_proc'):
            # For _proc columns: drop |values| > 1000
            numeric_df[col] = numeric_df[col].where(abs(numeric_df[col]) <= 1000)
        else:
            # For non-_proc columns: drop negatives
            numeric_df[col] = numeric_df[col].where(numeric_df[col] >= 0)
        # Round all values to 2 decimals
    numeric_df = numeric_df.round(2)
    return numeric_df

NaujiStulpeliai = {
    'IMONES_AMZIUS': 'Įmonės amžius (metai)',
    'VID_DARBUOTOJU_SK': 'Vidutinis darbuotojų skaičius',
    'VID_DARBUOTOJU_SK_POKYTIS_proc': 'Vidutinio darbuotojų skaičiaus pokytis (%)',
    'VID_Darbo_U': 'Vidutinis darbo užmokestis',
    'VID_Darbo_U_POKYTIS_proc': 'Vidutinio darbo užmokesčio pokytis (%)',
    'Pardavimai_POKYTIS_proc': 'Pardavimų pokytis (%)',
    'Pirkimai_POKYTIS_proc': 'Pirkimų pokytis (%)',
    'PARDAVIMO_PAJAMOS': 'Pardavimo pajamos',
    'NUOSAVAS_KAPITALAS': 'Nuosavas kapitalas',
    'Skola': 'Skola',
    'Bankroto_IVERTIS': 'Bankroto rizikos įvertis'
}

FilteredStatybos = filter_data(Statybos)
FilteredStatybos.rename(columns=NaujiStulpeliai, inplace=True)

FilteredPrekyba = filter_data(Prekyba)
FilteredPrekyba.rename(columns=NaujiStulpeliai, inplace=True)


# Statistical summary of the numeric columns

sns.set_style("whitegrid")

# Define a function to format the text and plot
def plot_distribution(df, column):
    plt.figure(figsize=(12, 8))

    # Histogram
    sns.histplot(df[column], color='green', stat='percent',kde=True, alpha=0.8, label='Kiekis (%)')


    # Mean line
    mean_value = df[column].mean()
    plt.axvline(mean_value, color='red', linestyle='--', label=f'Vidurkis: {mean_value:.2f}')

    # Title and labels
    plt.title(f'{column}, distribucija', fontsize=20, fontweight='bold', color='darkblue')
    plt.xlabel(column, fontsize=16, color='darkblue')
    plt.ylabel('Procentinė dalis (%)', fontsize=16, color='darkblue')

    # Legend
    plt.legend(fontsize=14, loc='upper right')

    # Grid lines
    plt.grid(True, which='both', linestyle='--', linewidth=0.5, color='gray')

    # Customize ticks
    plt.xticks(fontsize=14, color='darkblue')
    plt.yticks(fontsize=14, color='darkblue')

    # Tight layout to ensure everything fits well
    plt.tight_layout()

    # Save plot as PNG file in the same directory as the script
    file_name = f'{column}_distribucija.png'  # You can change the file extension if needed
    if os.path.exists(file_name):
        print(f"File '{file_name}' already exists. Skipping save.")
    else:
        plt.savefig(file_name, dpi=300, bbox_inches='tight')
        print(f"Plot saved as '{file_name}' with high quality.")
    # Show plot
    plt.show()


# Plot distribution for each column
for column in FilteredStatybos.columns:
    plot_distribution(FilteredStatybos, column)
for column in FilteredPrekyba.columns:
    plot_distribution(FilteredStatybos, column)