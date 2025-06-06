from pathlib import Path
import pandas as pd
import unicodedata
from rapidfuzz import process, fuzz

# Load your mapping from Excel file
def load_account_mapping(filepath=None, sheet_name=0):
    if filepath is None:
        # Default to a relative path in the project directory
        filepath = Path(__file__).parent / "Data" / "account_mapping.xlsx"
    else:
        filepath = Path(filepath)
    mapping_df = pd.read_excel(filepath, sheet_name=sheet_name)
    # Build mapping dictionary
    standard_account_map = {}
    for _, row in mapping_df.iterrows():
        std = str(row['standard_account']).strip()
        var = str(row['variant_account']).strip()
        if std in standard_account_map:
            standard_account_map[std].append(var)
        else:
            standard_account_map[std] = [var]
    return standard_account_map

# Normalize account names
def normalize_account(s):
    s = str(s).lower().strip()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    s = ' '.join(s.split())
    return s

# Build reverse mapping and fuzzy lists
def build_reverse_maps(standard_account_map):
    reverse_map = {}
    for std, variants in standard_account_map.items():
        for v in variants:
            reverse_map[normalize_account(v)] = std
    all_std_normalized = [normalize_account(k) for k in standard_account_map.keys()]
    all_variant_normalized = list(reverse_map.keys())
    return reverse_map, all_std_normalized, all_variant_normalized

# Robust lookup function using reverse map, with optional fuzzy matching
def robust_get(account, year, data_dict, reverse_map, all_std_normalized, all_variant_normalized, fuzzy_threshold=85):
    norm_account = normalize_account(account)
    # 1. Try direct match (normalized)
    for key in data_dict.keys():
        if normalize_account(key) == norm_account:
            return data_dict[key].get(str(year), None)
    # 2. Try mapping: if input is a variant, get standard; if input is standard, get all its variants
    mapped_standard = reverse_map.get(norm_account, None)
    if mapped_standard:
        # Try to find the standard account in data_dict
        for key in data_dict.keys():
            if normalize_account(key) == normalize_account(mapped_standard):
                return data_dict[key].get(str(year), None)
        # Try to find any variant in data_dict
        for key in data_dict.keys():
            if normalize_account(key) in all_variant_normalized and reverse_map[normalize_account(key)] == mapped_standard:
                return data_dict[key].get(str(year), None)
    else:
        # If input is a standard account, try all its variants
        for variant, std in reverse_map.items():
            if std == account or normalize_account(std) == norm_account:
                for key in data_dict.keys():
                    if normalize_account(key) == variant:
                        return data_dict[key].get(str(year), None)
    # 3. Fuzzy match as a last resort
    all_full_accounts = list(data_dict.keys())
    match, score, idx = process.extractOne(norm_account, [normalize_account(a) for a in all_full_accounts], scorer=fuzz.ratio)
    if score >= fuzzy_threshold:
        full_account = all_full_accounts[idx]
        return data_dict[full_account].get(str(year), None)
    # 4. Fallback: None
    return None

# Utility: For simple import in your main script
def setup_account_mapping(filepath='account_mapping.xlsx', sheet_name=0):
    standard_account_map = load_account_mapping(filepath, sheet_name)
    reverse_map, all_std_normalized, all_variant_normalized = build_reverse_maps(standard_account_map)
    return standard_account_map, reverse_map, all_std_normalized, all_variant_normalized
