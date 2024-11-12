import pandas as pd
import my_mods # type: ignore

# Set file paths and sheet names
file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/1Power Bi/Link Deviations.xlsx'
output_file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/POP (Play or Pay)/The Hub/Central Hub.xlsx'
city_id_file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/POP (Play or Pay)/CITY ID.xlsx'
pop_review_file_path = r'C:/Users/dsamu/dsamllc.net/dsamllc.net - Documents/FIS Project Documents/POP (Play or Pay)/POP review spreadsheet.xlsx'
PayrollInfo = r'C:\\Users\\dsamu\\dsamllc.net\\dsamllc.net - Documents\\FIS Project Documents\\1Power Bi\\PayrollInfo.xlsx'

HOT_LINK_sheet_name = 'HOT LINK'
ci_hub = 'City Id Hub'
pop_POPTrackingWorkBook = 'POPTrackingWorkBook'
sheet_pop = 'POP numbers'

ci_column = 'City Id'
contract_column = 'Contract Amount'
dates_au = 'oldest date'
weekly_au = 'weekly'
Q_au = 'quarterly'
sub_contractor_column_city_id = 'Sub-Contractor'
deviated_as_column = 'DEVIATED AS'
POP_3_Classification = 'POP 3 Classification'

# Step 1: Extract unique City Ids and write to Excel
unique_city_ids_df = my_mods.extract_unique_city_ids(file_path, HOT_LINK_sheet_name, ci_column)
if unique_city_ids_df is not None:
    my_mods.write_to_excel(unique_city_ids_df, output_file_path, 'POP Hub', mode='w')

# Step 2: Load City IDs and merge with unique City Ids using CityIdFinder class
city_id_finder = my_mods.CityIdFinder(city_id_file_path, ci_hub)
city_id_finder.load_city_id_data()

if city_id_finder.city_id_df is not None:
    merge_columns = [sub_contractor_column_city_id, deviated_as_column, POP_3_Classification]
    merged_df = my_mods.merge_city_dataframes(unique_city_ids_df, city_id_finder.city_id_df, ci_column, merge_columns)
    if merged_df is not None:
        my_mods.write_to_excel(merged_df, output_file_path, 'POP Hub', mode='a')

# Step 3: Sum Contract Amounts from the POP review file
try:
    pop_review_df = pd.read_excel(pop_review_file_path, sheet_name=pop_POPTrackingWorkBook)
    summed_contracts_df = my_mods.sum_contract_amounts(merged_df, pop_review_df, ci_column, contract_column)
    if summed_contracts_df is not None:
        final_df = pd.merge(merged_df, summed_contracts_df, on=ci_column, how='left')
except Exception as e:
    print(f"Error in Step 3: {e}")

# Step 4: Add POP Eligible column and save final DataFrame
try:
    final_df = my_mods.add_pop_eligible_column(final_df, contract_column, deviated_as_column)
    my_mods.write_to_excel(final_df, output_file_path, 'POP Hub', mode='a')
except Exception as e:
    print(f"Error in Step 4: {e}")

# Step 5-7: Append additional found information from POP review file for 'oldest date', 'weekly', and 'quarterly'
def append_information_by_column(output_file_path, pop_review_file_path, sheet_pop, hub_df, search_column):
    dfsheet_pop = pd.read_excel(pop_review_file_path, sheet_name=sheet_pop)
    fvic_df = my_mods.find_value_in_column(dfsheet_pop, ci_column, hub_df[ci_column].dropna().tolist(), search_column)
    if isinstance(fvic_df, pd.DataFrame):
        combined_df = my_mods.append_found_information(hub_df, fvic_df, ci_column)
        my_mods.write_to_excel(combined_df, output_file_path, 'POP Hub')

try:
    hub_df = pd.read_excel(output_file_path, sheet_name='POP Hub')

    append_information_by_column(output_file_path, pop_review_file_path, sheet_pop, hub_df, dates_au)
except Exception as e:
    print(f"Error in Step 5: {e}")
try:
    hub_df = pd.read_excel(output_file_path, sheet_name='POP Hub')
    append_information_by_column(output_file_path, pop_review_file_path, sheet_pop, hub_df, weekly_au)
    print(123)
except Exception as e:
    print(f"Error in Step 6: {e}")

try:
    hub_df = pd.read_excel(output_file_path, sheet_name='POP Hub')
    append_information_by_column(output_file_path, pop_review_file_path, sheet_pop, hub_df, Q_au)
except Exception as e:
    print(f"Error in Step 7: {e}")



