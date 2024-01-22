import pandas as pd
import scipy.stats as stats


def load_and_preprocess(path):
        data = pd.read_excel(path, sheet_name='TABLICA', header=0, index_col=0)
        return data.apply(pd.to_numeric, errors='coerce').dropna(how='all', axis=1)

def calculate_correlation(data_1_unit, data_2_unit, correlation_type):
    if correlation_type == 'pearson':
        corr, p_value = stats.pearsonr(data_1_unit, data_2_unit)
    elif correlation_type == 'spearman':
        corr, p_value = stats.spearmanr(data_1_unit, data_2_unit)
    elif correlation_type == 'kendall':
        corr, p_value = stats.kendalltau(data_1_unit, data_2_unit)
    return corr, p_value

def statistic_test(data_1_file_name, data_2_file_name, correlation_type = 'pearson'):
    # File paths for the latest uploaded files
    path_1 = f'Data/{data_1_file_name}.xlsx'
    path_2 = f'Data/{data_2_file_name}.xlsx'

    # Load and preprocess data
    data_1 = load_and_preprocess(path_1)
    data_2 = load_and_preprocess(path_2)

    # Initialize list for new rows
    new_rows = []

    # Calculate correlation and p-value
    for unit in data_1.index.intersection(data_2.index):
        data_1_unit = data_1.loc[unit].dropna()
        data_2_unit = data_2.loc[unit].dropna()

        if len(data_1_unit) > 1 and len(data_2_unit) > 1:
            n = len(data_1_unit)
            corr, p_value = calculate_correlation(data_1_unit, data_2_unit, correlation_type)
            t_statistic = corr * ((n - 2) ** 0.5) / ((1 - corr ** 2) ** 0.5)
            degrees_of_freedom = n - 2
            significance = p_value < 0.05
            new_rows.append({
                'Administrative_Unit': unit,
                'Correlation': corr, 
                'P-Value': p_value, 
                'T-Statistic': t_statistic, 
                'Degrees of Freedom': degrees_of_freedom,
                'Significance': significance
            })

    # Create DataFrame from new rows
    correlation_results = pd.DataFrame(new_rows)

    # Save the results
    output_file_path = f"Results/Test_{data_1_file_name}_{data_2_file_name}_{correlation_type}.xlsx"
    correlation_results.to_excel(output_file_path, index=False)


def statistic_test_3(data_1_file_name, data_2_file_name, data_3_file_name, correlation_type = 'pearson'):
    # File paths for the latest uploaded files
    path_1 = f'Data\\{data_1_file_name}.xlsx'
    path_2 = f'Data\\{data_2_file_name}.xlsx'
    path_3 = f'Data\\{data_3_file_name}.xlsx'

    # Load and preprocess data
    data_1 = load_and_preprocess(path_1)
    data_2 = load_and_preprocess(path_2)
    data_3 = load_and_preprocess(path_3)

    # Initialize list for new rows
    new_rows = []

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
            corr_1_3, p_value_1_3 = calculate_correlation(data_1_unit_aligned, data_3_unit_aligned, correlation_type)
            corr_2_3, p_value_2_3 = calculate_correlation(data_2_unit_aligned, data_3_unit_aligned, correlation_type)
            new_rows.append({
                'Administrative_Unit': unit, 
                'Correlation_1_3': corr_1_3, 
                'Correlation_2_3': corr_2_3, 
                'P-Value_1_3': p_value_1_3, 
                'P-Value_2_3': p_value_2_3,
                'Significance_1_3': p_value_1_3 < 0.05,
                'Significance_2_3': p_value_2_3 < 0.05,
            })
    
    # Create DataFrame from new rows
    correlation_results = pd.DataFrame(new_rows)

    # Save the results
    output_file_path = f"Results\\Test_{data_1_file_name}_{data_2_file_name}_{data_3_file_name}_{correlation_type}.xlsx"
    correlation_results.to_excel(output_file_path, index=False)


# File names
file_name_zaniesczyszczenia = ['Powietrze_99', 'Powietrze_03', 'Powietrze_04']
file_name_zdrowie = 'Gruzlica_99'
file_name_ludnosc = 'Ludnosc_99'
file_name_pojazdy = ['Samochody_ciezar_03', 'Samochody_ciezar_04']
file_name_drogi = 'Drogi_04'

# Correlation types
test_types = ['pearson', 'spearman', 'kendall']

# Test pairs
file_name_test_pairs = [
    [file_name_zaniesczyszczenia[0], file_name_zdrowie],
    [file_name_zaniesczyszczenia[0], file_name_ludnosc],
    [file_name_zaniesczyszczenia[1], file_name_pojazdy[0]],
    [file_name_zaniesczyszczenia[2], file_name_drogi]
]

# Test trios
file_name_test_trios = [
    [file_name_zaniesczyszczenia[0], file_name_zdrowie, file_name_ludnosc],
    [file_name_zaniesczyszczenia[2], file_name_drogi, file_name_pojazdy[1]]
]


'''# Test for pairs
for test_pair in file_name_test_pairs:
    for test_type in test_types:
        statistic_test(test_pair[0], test_pair[1], test_type)'''

# Test for trios
for test_trio in file_name_test_trios:
    for test_type in test_types:
        statistic_test_3(test_trio[0], test_trio[1], test_trio[2], test_type)
