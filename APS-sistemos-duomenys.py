import os
import textwrap
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

print("All libraries imported successfully!")



# Replace 'your_file.xlsx' with the path to your Excel file
df = pd.read_excel('APS_dalyviai_duomenys iš VMI.xlsx', header=0)
df = df.dropna(how='all')

def is_all_strings(row):
    return all(isinstance(item, str) for item in row)

# Filter out rows where all values are strings
df = df[~df.apply(is_all_strings, axis=1)]

# Reset the index after filtering
df.reset_index(drop=True, inplace=True)

print(df.head())

# Split the DataFrame into four parts, each with 201 rows
q1 = df.iloc[:200]
q2 = df.iloc[200:400]
q3 = df.iloc[400:600]
q4 = df.iloc[601:800]

print(q1.head())
print(q2.head())
print(q3.head())
print(q4.head())

AllQuarters = pd.concat([q1, q2, q3, q4])

#Vidutiniai imoniu amžiai
avgQ1companyAge = q1['IMONES_AMZIUS'].mean()
avgQ2companyAge = q2['IMONES_AMZIUS'].mean()
avgQ3companyAge = q3['IMONES_AMZIUS'].mean()
avgQ4companyAge = q4['IMONES_AMZIUS'].mean()
avg2024companyAge = AllQuarters['IMONES_AMZIUS'].mean()

#Vidutinis darbuotojų skaičius
avgQ1numberOfEmployees = q1['VID_DARBUOTOJU_SK'].mean()
avgQ2numberOfEmployees = q2['VID_DARBUOTOJU_SK'].mean()
avgQ3numberOfEmployees = q3['VID_DARBUOTOJU_SK'].mean()
avgQ4numberOfEmployees = q4['VID_DARBUOTOJU_SK'].mean()
avg2024numberOfEmployees = AllQuarters['VID_DARBUOTOJU_SK'].mean()

#Klasifikatoriu reiksmes

# Data
data = {
    'code': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U'],
    'meaning': [
        'AGRICULTURE, FORESTRY AND FISHING',
        'MINING AND QUARRYING',
        'MANUFACTURING',
        'ELECTRICITY, GAS, STEAM AND AIR CONDITIONING SUPPLY',
        'WATER SUPPLY; SEWERAGE, WASTE MANAGEMENT AND REMEDIATION ACTIVITIES',
        'CONSTRUCTION',
        'WHOLESALE AND RETAIL TRADE; REPAIR OF MOTOR VEHICLES AND MOTORCYCLE',
        'TRANSPORTATION AND STORAGE',
        'ACCOMMODATION AND FOOD SERVICE ACTIVITIES',
        'INFORMATION AND COMMUNICATION',
        'FINANCIAL AND INSURANCE ACTIVITIES',
        'REAL ESTATE ACTIVITIES',
        'PROFESSIONAL, SCIENTIFIC AND TECHNICAL ACTIVITIES',
        'ADMINISTRATIVE AND SUPPORT SERVICE ACTIVITIES',
        'PUBLIC ADMINISTRATION AND DEFENCE; COMPULSORY SOCIAL SECURITY',
        'EDUCATION',
        'HUMAN HEALTH AND SOCIAL WORK ACTIVITIES',
        'ARTS, ENTERTAINMENT AND RECREATION',
        'OTHER SERVICE ACTIVITIES',
        'ACTIVITIES OF HOUSEHOLDS AS EMPLOYERS; UNDIFFERENTIATED GOODS- AND SERVICES-PRODUCING ACTIVITIES OF HOUSEHOLDS FOR OWN USE',
        'ACTIVITIES OF EXTRATERRITORIAL ORGANISATIONS AND BODIES'
    ]
}

lentele = pd.DataFrame(data)


def count_and_filter(df, column_name):
    # Initialize a dictionary to store the counts
    counts = {chr(i): 0 for i in range(ord('A'), ord('U') + 1)}

    # Count the number of elements starting with each letter from A to U
    for letter in counts.keys():
        counts[letter] = df[column_name].str.startswith(letter).sum()

    # Convert the dictionary to a DataFrame for better readability
    counts_df = pd.DataFrame(list(counts.items()), columns=['Code', 'Count'])

    # Remove rows with Count equal to 0
    counts_df = counts_df[counts_df['Count'] != 0]

    return counts_df

#I APS patekusiu imoniu veiklos sferos skirtinguose ketvircuose:

sectorsQ1 = count_and_filter(q1, 'EVRK')
sectorsQ2 = count_and_filter(q2, 'EVRK')
sectorsQ3 = count_and_filter(q3, 'EVRK')
sectorsQ4 = count_and_filter(q4, 'EVRK')
sectors2024 = count_and_filter(AllQuarters, 'EVRK')

# I APS patekusiu imoniu veiklos sferos skirtinguose ketvircuose, isrusiuotos nuo didziausiu:
sortedSectorsQ1 = sectorsQ1.sort_values('Count', ascending=False)
sortedSectorsQ2 = sectorsQ2.sort_values('Count', ascending=False)
sortedSectorsQ3 = sectorsQ3.sort_values('Count', ascending=False)
sortedSectorsQ4 = sectorsQ4.sort_values('Count', ascending=False)
sortedSectors2024 = sectors2024.sort_values('Count', ascending=False)

#Vizualizacija


def plot_sector_distribution(df, title):

    plt.figure(figsize=(16, 12))  # 4:3 aspect ratio

    # Calculate percentages
    total = df['Count'].sum()
    df['Percentage'] = (df['Count'] / total) * 100
    df = df.sort_values('Percentage', ascending=False)

    # Create the plot
    ax = sns.barplot(x='Percentage', y='Code', data=df,
                     palette='viridis',
                     edgecolor='black',
                     linewidth=0.8)

    # Add percentage labels
    for p in ax.patches:
        width = p.get_width()
        if width > 3:
            ax.text(width - 1,
                    p.get_y() + p.get_height() / 2.,
                    f'{width:.1f}%',
                    ha='right', va='center',
                    fontsize=10, weight='bold', color='white')
        else:
            ax.text(width + 0.3,
                    p.get_y() + p.get_height() / 2.,
                    f'{width:.1f}%',
                    ha='left', va='center',
                    fontsize=10, weight='bold')

    # Sector mapping dictionary
    sector_dict = {
        'A': 'Žemės ūkis, miškininkystė ir žuvininkystė',
        'B': 'Kasyba ir karjerų eksploatavimas',
        'C': 'Apdirbamoji gamyba',
        'D': 'Elektros, dujų, garo tiekimas ir oro kondicionavimas',
        'E': 'Vandens tiekimas, nuotekų valymas, atliekų tvarkymas',
        'F': 'Statyba',
        'G': 'Didmeninė ir mažmeninė prekyba; transporto priemonių remontas',
        'H': 'Transportas ir saugojimas',
        'I': 'Apgyvendinimo ir maitinimo paslaugų veikla',
        'J': 'Informacija ir ryšiai',
        'K': 'Finansinė ir draudimo veikla',
        'L': 'Nekilnojamojo turto operacijos',
        'M': 'Profesinė, mokslinė ir techninė veikla',
        'N': 'Administracinė ir aptarnavimo veikla',
        'O': 'Viešasis valdymas ir gynyba; socialinis draudimas',
        'P': 'Švietimas',
        'Q': 'Žmonių sveikatos priežiūra ir socialinis darbas',
        'R': 'Meninė, pramoginė ir poilsio organizavimo veikla',
        'S': 'Kita aptarnavimo veikla',
        'T': 'Namų ūkių veikla',
        'U': 'Ekstrateritorinių organizacijų veikla'
    }

    # Text wrapping function
    def wrap_text(text, width=35):
        return '\n'.join(textwrap.wrap(text, width=width))

    # Create legend with wrapped text
    present_sectors = df['Code'].unique()
    legend_text = [f"$\mathbf{{ {code} }}$: {wrap_text(sector_dict[code])}"
                   for code in present_sectors]

    legend = plt.legend(handles=[plt.Line2D([0], [0], marker='', color='w', label=text,
                                            markersize=0) for text in legend_text],
                        bbox_to_anchor=(1.02, 1),
                        loc='upper left',
                        borderaxespad=0.,
                        title="EVRK sektorių legenda",
                        title_fontsize='12',
                        fontsize='9.5',
                        frameon=True,
                        facecolor='#f8f9fa',
                        edgecolor='#dee2e6',
                        shadow=False,
                        fancybox=True,
                        labelspacing=0.8,
                        handlelength=0,
                        handletextpad=0)

    legend._legend_box.align = "left"
    plt.setp(legend.get_texts(), linespacing=1.4)

    # Add titles and labels
    plt.title(title,
              fontsize=16, pad=20, weight='bold', color='#2c3e50')
    plt.xlabel('Procentinė dalis (%)', fontsize=12, labelpad=10, weight='bold')
    plt.ylabel('EVRK kodas', fontsize=12, labelpad=10, weight='bold')

    # Customize ticks and grid
    plt.xticks(fontsize=11)
    plt.yticks(fontsize=11)
    plt.grid(axis='x', linestyle='--', alpha=0.3, color='gray')

    # Final adjustments
    sns.despine(left=True, bottom=True)
    plt.tight_layout(rect=[0, 0, 0.72, 0.95])
    plt.subplots_adjust(right=0.68)

    # Convert title to a valid file name
    filename = ''.join(e for e in title if e.isalnum() or e in ' .-').strip() + '.pdf'

    # Save the plot with the title as the file name


    if os.path.exists(filename):
        print(f"File '{filename}' already exists. Skipping save.")
    else:
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        print(f"Plot saved as '{filename}' with high quality.")
    plt.show()
# I APS patekusiu imoniu veiklos sferos skirtinguose ketvircuose, isrusiuotos nuo didziausiu (%):
q1SectorPercentagePlot =plot_sector_distribution(sortedSectorsQ1, 'Į APS patekusių įmonių veiklos sferos I ketvirtyje')
q2SectorPercentagePlot =plot_sector_distribution(sortedSectorsQ2, 'Į APS patekusių įmonių veiklos sferos II ketvirtyje')
q3SectorPercentagePlot =plot_sector_distribution(sortedSectorsQ3, 'Į APS patekusių įmonių veiklos sferos III ketvirtyje')
q4SectorPercentagePlot =plot_sector_distribution(sortedSectorsQ4, 'Į APS patekusių įmonių veiklos sferos IV ketvirtyje')
allQuartersSectorPercentagePlot =plot_sector_distribution(sortedSectors2024, 'Į APS patekusių įmonių veiklos sferos')

#------------------------------------------------------------------------------------------------------------------------------------