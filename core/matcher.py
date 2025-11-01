import re
import pandas as pd
from difflib import SequenceMatcher
import logging

def normalize_name(name):
    if pd.isna(name):
        return ""
    cleaned = re.sub(r"^(Dr\.?|Prof\.?|Mr\.?|Mrs\.?|Ms\.?)\s*", "", str(name).strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"\.(?=\w)", " ", cleaned)
    parts = [part.lower() for part in cleaned.split() if part]
    return ' '.join(parts)

def normalize_designation(desig):
    if pd.isna(desig):
        return ""
    desig = str(desig).strip().lower()
    designation_map = {
        'professor': 'Professor', 'prof': 'Professor',
        'assoc. professor': 'Assoc. Professor', 'asp': 'Assoc. Professor', 'associate prof': 'Assoc. Professor',
        'asst. professor': 'Asst. Professor', 'ap': 'Asst. Professor', 'asst prof': 'Asst. Professor',
        'a.p(contract)': 'A.P(Contract)', 'gl': 'A.P(Contract)', 'guest lecturer': 'A.P(Contract)'
    }
    return designation_map.get(desig, desig.title())

def fuzzy_match_name(staff_name, pref_name, threshold=0.65):
    try:
        staff_norm = normalize_name(staff_name)
        pref_norm = normalize_name(pref_name)
        staff_parts = staff_norm.split()
        pref_parts = pref_norm.split()
        
        common_parts = set(staff_parts) & set(pref_parts)
        if any(len(part) > 3 for part in common_parts):
            logging.info(f"Fuzzy matched {pref_name} to {staff_name} based on significant part: {common_parts}")
            return True
        
        score = SequenceMatcher(None, staff_norm, pref_norm).ratio()
        logging.debug(f"Fuzzy match attempt: {staff_name} vs {pref_name}, normalized: {staff_norm} vs {pref_norm}, score: {score:.3f}")
        
        staff_raw = re.sub(r"[.\s]+", "", staff_name.lower())
        pref_raw = re.sub(r"[.\s]+", "", pref_name.lower())
        raw_score = SequenceMatcher(None, staff_raw, pref_raw).ratio()
        logging.debug(f"Raw match attempt: {staff_raw} vs {pref_raw}, score: {raw_score:.3f}")
        
        if score >= threshold or raw_score >= 0.9:
            logging.info(f"Fuzzy matched {pref_name} to {staff_name} with score {max(score, raw_score):.3f}")
            return True
        
        logging.debug(f"No match for {staff_name} vs {pref_name}, scores: normalized={score:.3f}, raw={raw_score:.3f}, threshold={threshold}")
        return False
    except Exception as e:
        logging.error(f"Fuzzy match failed for {staff_name} vs {pref_name}: {e}")
        return False

