from scipy.stats import norm
import numpy as np
from fuzzywuzzy import fuzz


def data_frame_to_np(df):
    column_names = df.columns.to_numpy()
    data_array = df.to_numpy()
    combined_array = np.vstack([column_names, data_array])
    return combined_array


def calculate_string_similarity(str1, str2):
    ratio_similarity = fuzz.ratio(str1, str2)
    partial_ratio_similarity = fuzz.partial_ratio(str1, str2)
    token_sort_similarity = fuzz.token_sort_ratio(str1, str2)
    return max(ratio_similarity, partial_ratio_similarity, token_sort_similarity)


def print_red(text):
    print("\033[31m" + text + "\033[0m")


def normalize_ratings(data, desired_mean=5.0, value_range=(1, 10)):
    data = np.array(data)
    min_val, max_val = value_range

    current_mean = data.mean()

    shift_value = desired_mean - current_mean
    adjusted = data + shift_value

    adjusted = np.clip(adjusted, min_val, max_val)

    final_mean = np.mean(adjusted)
    if final_mean != desired_mean:
        fine_tune_diff = desired_mean - final_mean
        adjusted += fine_tune_diff
        adjusted = np.clip(adjusted, min_val, max_val)
    adjusted = adjusted.tolist()
    adjusted = [round(x, 3) for x in adjusted]

    return adjusted

def find_placement_score(rank, num_teams=60, multiplier=30, mean=.5704, std_dev=.22,
                         min_percentile=0.01,
                         max_percentile=0.99):
    """
    Generalized difficulty score using rank, with safe percentile bounds to avoid infinities.
    """
    rank = min(max(rank, 1), num_teams)

    # Normalize to [0, 1], then scale to [min_percentile, max_percentile]
    percentile = max_percentile - ((rank - 1) / (num_teams - 1)) * (max_percentile - min_percentile)
    percentile = min(max(percentile, min_percentile), max_percentile)  # extra safety

    z = norm.ppf(percentile)
    output = mean + z * std_dev * multiplier
    return round(output, 2)


def reverse_placement_score(output, num_teams=60, multiplier=30, mean=0.5704, std_dev=0.22,
                            min_percentile=0.01, max_percentile=0.99):
    # Step 1: Compute z-score
    z = (output - mean) / (std_dev * multiplier)

    # Step 2: Convert z-score to percentile
    percentile = norm.cdf(z)

    # Clip to safe percentile bounds
    percentile = min(max(percentile, min_percentile), max_percentile)

    # Step 3: Convert percentile to rank
    rank = 1 + (max_percentile - percentile) / (max_percentile - min_percentile) * (num_teams - 1)

    return round(rank)


