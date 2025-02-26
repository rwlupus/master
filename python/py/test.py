import pandas as pd
import numpy as np

def compare_column_changes(curr_df, prev_df, key_column, numeric_columns, custom_thresholds=None, default_threshold=None):
    if default_threshold is None:
        default_threshold = {"min": -0.1, "max": 0.1, "use_percentage": True}
    if custom_thresholds is None:
        custom_thresholds = {}

    # Merge danych na podstawie klucza
    merged = prev_df.merge(curr_df, on=key_column, suffixes=("_prev", "_curr"), how="outer")

    results = []
    
    for col in numeric_columns:
        prev_col = f"{col}_prev"
        curr_col = f"{col}_curr"

        # Pobieramy thresholdy dla danej kolumny (jeśli brak – używamy domyślnych)
        thresholds = custom_thresholds.get(col, default_threshold)
        min_thresh = thresholds["min"]
        max_thresh = thresholds["max"]
        use_percentage = thresholds["use_percentage"]

        # Tworzymy tymczasowe kolumny do obliczeń (nie zmieniamy oryginałów)
        prev_clean = merged[prev_col].replace("", np.nan).fillna(0)
        curr_clean = merged[curr_col].replace("", np.nan).fillna(0)

        # Sprawdzenie, które wartości są numeryczne
        is_prev_numeric = pd.to_numeric(prev_clean, errors="coerce").notna()
        is_curr_numeric = pd.to_numeric(curr_clean, errors="coerce").notna()

        # Obliczamy różnicę absolutną (tylko jeśli obie wartości są numeryczne)
        merged["difference"] = np.where(is_prev_numeric & is_curr_numeric, curr_clean - prev_clean, np.nan)

        # Obliczamy różnicę procentową, zabezpieczając przed dzieleniem przez zero
        merged["difference %"] = np.where(
            (prev_clean != 0) & is_prev_numeric & is_curr_numeric,
            ((curr_clean - prev_clean) / prev_clean) * 100,
            np.nan
        )

        # Określamy, czy różnica przekroczyła próg
        if use_percentage:
            merged["Threshold_exceeded"] = (merged["difference %"] < min_thresh) | (merged["difference %"] > max_thresh)
        else:
            merged["Threshold_exceeded"] = (merged["difference"] < min_thresh) | (merged["difference"] > max_thresh)

        # Oznaczanie dodatkowych sytuacji:
        # 1. Jeśli wartość prev była 0, a curr jest liczbą -> oznaczamy jako przekroczony próg
        merged["Threshold_exceeded"] |= (prev_clean == 0) & is_curr_numeric & (curr_clean != 0)

        # 2. Jeśli którakolwiek wartość jest nienumeryczna -> oznaczamy jako przekroczony próg
        merged["Threshold_exceeded"] |= (~is_prev_numeric | ~is_curr_numeric)

        # Dodajemy kolumnę określającą, czy próg był liczony w procentach
        merged["Threshold%"] = use_percentage

        # Wybieramy tylko istotne kolumny i dodajemy do wyników
        col_results = merged.loc[merged["Threshold_exceeded"], [key_column, "difference", "difference %", "Threshold%", "Threshold_exceeded"]]
        col_results.insert(1, "column", col)  # Dodajemy nazwę kolumny jako nową kolumnę

        results.append(col_results)

    # Łączymy wyniki dla wszystkich kolumn w jeden dataframe
    final_results = pd.concat(results, ignore_index=True)

    # Sortujemy wyniki według key_column i column
    final_results = final_results.sort_values(by=[key_column, "column"]).reset_index(drop=True)

    return final_results
