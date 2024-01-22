import pandas as pd
import scipy.stats as stats

def statistic_test(data_1_file_name, data_2_file_name):
    # File paths for the latest uploaded files
    path_1 = f'Data\\{data_1_file_name}.xlsx'
    path_2 = f'Data\\{data_2_file_name}.xlsx'

    # Load the data from the 'TABLICA' sheet
    data_1 = pd.read_excel(path_1, sheet_name='TABLICA', header=0, index_col=0)
    data_2 = pd.read_excel(path_2, sheet_name='TABLICA', header=0, index_col=0)

    # Convert data to numeric and drop rows/columns with all NaN values
    data_1 = data_1.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)
    data_2 = data_2.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)

    # Initialize the DataFrame with the correct column types
    correlation_results = pd.DataFrame(columns=['Administrative_Unit', 'Correlation', 'P-Value'])

    # Perform the loop and concatenation
    for unit in data_1.index.intersection(data_2.index):
        data_1_unit = data_1.loc[unit].dropna()
        data_2_unit = data_2.loc[unit].dropna()

        if len(data_1_unit) > 1 and len(data_2_unit) > 1:
            corr, p_value = stats.pearsonr(data_1_unit, data_2_unit)
            new_row = pd.DataFrame({'Administrative_Unit': [unit], 'Correlation': [corr], 'P-Value': [p_value]})
            correlation_results = pd.concat([correlation_results, new_row], ignore_index=True)

    # Performing statistical tests
    correlation_results['T-Statistic'] = None
    correlation_results['Degrees of Freedom'] = None
    correlation_results['Significance'] = None

    for index, row in correlation_results.iterrows():
        unit = row['Administrative_Unit']
        data_1_unit = data_1.loc[unit].dropna()
        data_2_unit = data_2.loc[unit].dropna()

        n = len(data_1_unit)
        r = row['Correlation']
        t_statistic = r * ((n - 2) ** 0.5) / ((1 - r ** 2) ** 0.5)
        degrees_of_freedom = n - 2
        p_value = 2 * (1 - stats.t.cdf(abs(t_statistic), df=degrees_of_freedom))

        correlation_results.at[index, 'T-Statistic'] = t_statistic
        correlation_results.at[index, 'Degrees of Freedom'] = degrees_of_freedom
        correlation_results.at[index, 'Significance'] = p_value < 0.05

    # Print and save the results
    print(correlation_results)
    output_file_path = f"Results\\Test_{data_1_file_name}_{data_2_file_name}.xlsx"
    correlation_results.to_excel(output_file_path, index=False)

def statistic_test_3(data_1_file_name, data_2_file_name, data_3_file_name):
    # File paths for the latest uploaded files
    path_1 = f'Data\\{data_1_file_name}.xlsx'
    path_2 = f'Data\\{data_2_file_name}.xlsx'
    path_3 = f'Data\\{data_3_file_name}.xlsx'

    # Load the data from the 'TABLICA' sheet
    data_1 = pd.read_excel(path_1, sheet_name='TABLICA', header=0, index_col=0)
    data_2 = pd.read_excel(path_2, sheet_name='TABLICA', header=0, index_col=0)
    data_3 = pd.read_excel(path_3, sheet_name='TABLICA', header=0, index_col=0)  # Corrected from path_2 to path_3

    # Convert data to numeric and drop rows/columns with all NaN values
    data_1 = data_1.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)
    data_2 = data_2.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)
    data_3 = data_3.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)

    # Initialize the DataFrame with the correct column types
    correlation_results = pd.DataFrame(columns=['Administrative_Unit', 'Correlation_1_3', 'Correlation_2_3', 'P-Value_1_3', 'P-Value_2_3'])

    # Calculate the correlations
    for unit in set(data_1.index) & set(data_2.index) & set(data_3.index):
        data_1_unit = data_1.loc[unit].dropna()
        data_2_unit = data_2.loc[unit].dropna()
        data_3_unit = data_3.loc[unit].dropna()
    
        # Ensure all data are for the same period
        common_indices = list(set(data_1_unit.index) & set(data_2_unit.index) & set(data_3_unit.index))
        data_1_unit_aligned = data_1_unit.loc[common_indices]
        data_2_unit_aligned = data_2_unit.loc[common_indices]
        data_3_unit_aligned = data_3_unit.loc[common_indices]
        
        if len(data_1_unit_aligned) > 1 and len(data_2_unit_aligned) > 1:
            corr_1_3, p_value_1_3 = stats.pearsonr(data_1_unit_aligned, data_3_unit_aligned)
            corr_2_3, p_value_2_3 = stats.pearsonr(data_2_unit_aligned, data_3_unit_aligned)
            new_row = pd.DataFrame({
                'Administrative_Unit': [unit], 
                'Correlation_1_3': [corr_1_3], 
                'Correlation_2_3': [corr_2_3], 
                'P-Value_1_3': [p_value_1_3], 
                'P-Value_2_3': [p_value_2_3],
                'Significance_1_3': p_value_1_3 < 0.05,
                'Significance_2_3': p_value_2_3 < 0.05,
            })
            correlation_results = pd.concat([correlation_results, new_row], ignore_index=True)

    # No need to perform statistical tests again since p-values are already calculated by pearsonr function
    # Print and save the results
    print(correlation_results)
    output_file_path = f"Results\\Test_{data_1_file_name}_{data_2_file_name}_{data_3_file_name}.xlsx"
    correlation_results.to_excel(output_file_path, index=False)


# File names and test pairs
file_name_zaniesczyszczenia = ['Powietrze_99', 'Powietrze_03', 'Powietrze_04']
file_name_zdrowie = 'Gruzlica_99'
file_name_ludnosc = 'Ludnosc_99'
file_name_pojazdy = ['Samochody_ciezar_03', 'Samochody_ciezar_04']
file_name_drogi = 'Drogi_04'
        
file_name_test_pairs = [
    [file_name_zaniesczyszczenia[0], file_name_zdrowie],
    [file_name_zaniesczyszczenia[0], file_name_ludnosc],
    [file_name_zaniesczyszczenia[1], file_name_pojazdy[0]],
    [file_name_zaniesczyszczenia[2], file_name_drogi]
]

file_name_test_trios = [
    [file_name_zaniesczyszczenia[0], file_name_zdrowie, file_name_ludnosc],
    [file_name_zaniesczyszczenia[2], file_name_drogi, file_name_pojazdy[1]]
]

#for test_pair in file_name_test_pairs:
#    statistic_test(test_pair[0], test_pair[1])

for test_trio in file_name_test_trios:
    statistic_test_3(test_trio[0], test_trio[1], test_trio[2])


