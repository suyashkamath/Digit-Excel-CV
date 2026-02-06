# # #!/usr/bin/env python3
# # """
# # Unified Digit CV Processing API
# # ===============================
# # This is the FastAPI version of the original script, preserving all core logic.
# # It provides endpoints for getting sheet names and processing sheets.

# # Usage:
# #     uvicorn app:app --reload

# # Features:
# # - Upload Excel file to get sheet names
# # - Process specific sheet or all sheets if 'all' specified
# # - If no sheet_name provided, process the first sheet (preserving "process as is" logic)
# # - Returns processed records as JSON and output Excel as base64
# # - Automatic pattern detection and processing
# # """

# # import pandas as pd
# # import io
# # import re
# # import os
# # import sys
# # from typing import List, Dict, Tuple, Optional
# # from datetime import datetime
# # from fastapi import FastAPI, UploadFile, File, Form
# # from typing import Optional
# # import base64

# # app = FastAPI()

# # # ===============================================================================
# # # FORMULA DATA AND STATE MAPPING
# # # ===============================================================================

# # FORMULA_DATA = [
# #     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
# #     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
# #     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
# #     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
# # ]

# # STATE_MAPPING = {
# #     "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
# #     "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
# #     "Surat": "GUJARAT", "Jaipur": "RAJASTHAN", "Lucknow": "UTTAR PRADESH",
# #     "Patna": "BIHAR", "Ranchi": "JHARKHAND", "Bhuvaneshwar": "ODISHA",
# #     "Srinagar": "JAMMU AND KASHMIR", "Dehradun": "UTTARAKHAND", "Haridwar": "UTTARAKHAND",
# #     "Himachal Pradesh": "HIMACHAL PRADESH", "Andaman": "ANDAMAN AND NICOBAR ISLANDS",
# #     "Bangalore": "KARNATAKA", "Jharkhand": "JHARKHAND", "Bihar": "BIHAR",
# #     "West Bengal": "WEST BENGAL", "North Bengal": "WEST BENGAL", "Orissa": "ODISHA",
# #     "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
# #     "ROM2": "REST OF MAHARASHTRA", "Good Vizag": "ANDHRA PRADESH", "Good TN": "TAMIL NADU",
# #     "Kerala": "KERALA", "Good MP": "MADHYA PRADESH", "Good CG": "CHHATTISGARH",
# #     "Good RJ": "RAJASTHAN", "Bad RJ": "RAJASTHAN", "Good UP": "UTTAR PRADESH",
# #     "Bad UP": "UTTAR PRADESH", "Good UK": "UTTARAKHAND", "Bad UK": "UTTARAKHAND",
# #     "Punjab": "PUNJAB", "Jammu": "JAMMU AND KASHMIR", "Assam": "ASSAM",
# #     "NE EX ASSAM": "NORTH EAST", "Good NL": "NAGALAND", "GOOD KA": "KARNATAKA",
# #     "BAD KA": "KARNATAKA", "HR Ref": "HARYANA", "Dehradun, Haridwar": "UTTARAKHAND",
# #     "NE excl Assam": "NORTH EAST", "RJ REF": "RAJASTHAN"
# # }

# # # ===============================================================================
# # # CORE CALCULATION FUNCTIONS
# # # ===============================================================================

# # def get_payin_category(payin: float) -> str:
# #     """Categorize payin percentage into predefined ranges."""
# #     if payin <= 20:
# #         return "Payin Below 20%"
# #     elif payin <= 30:
# #         return "Payin 21% to 30%"
# #     elif payin <= 50:
# #         return "Payin 31% to 50%"
# #     else:
# #         return "Payin Above 50%"


# # def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
# #     """Calculate payout based on LOB, segment, policy type, and payin percentage."""
# #     if payin == 0:
# #         return 0, "0% (No Payin)", "Payin is 0, so Payout is 0"
    
# #     payin_category = get_payin_category(payin)
# #     matching_rule = None
    
# #     for rule in FORMULA_DATA:
# #         if rule["LOB"] == lob and rule["SEGMENT"] == segment:
# #             if rule["REMARKS"] == payin_category:
# #                 matching_rule = rule
# #                 break
    
# #     if not matching_rule:
# #         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
# #         payout = round(payin - deduction, 2)
# #         return payout, f"-{deduction}%", f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"
    
# #     formula = matching_rule["PO"]
# #     if formula.startswith("-") and "%" in formula:
# #         deduction = float(formula.replace("%", "").replace("-", ""))
# #         payout = round(payin - deduction, 2)
# #         return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"
# #     else:
# #         deduction = 2
# #         payout = round(payin - deduction, 2)
# #         return payout, f"-{deduction}%", f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"

# # # ===============================================================================
# # # VALUE EXTRACTION UTILITIES
# # # ===============================================================================

# # def safe_float(value) -> Optional[float]:
# #     """Safely convert value to float, handling various edge cases."""
# #     if pd.isna(value):
# #         return None
# #     val_str = str(value).strip().upper()
# #     if val_str in ["D", "NA", "", "NAN", "NONE"]:
# #         return None
# #     try:
# #         num = float(val_str.replace('%', '').strip())
# #         if 0 < num < 1:
# #             num = num * 100
# #         return num
# #     except:
# #         return None


# # def extract_lowest_payin(cell_value) -> Optional[float]:
# #     """
# #     Extract the lowest valid percentage from a cell that may contain multiple values.
# #     Examples: '49.5%/44.5%', '50 / 45.5', '48% - 52%'
# #     """
# #     if pd.isna(cell_value):
# #         return None
    
# #     cell_str = str(cell_value).strip()
# #     if not cell_str or cell_str.upper() in ["D", "NA", "NAN", "NONE", "", "D"]:
# #         return None
    
# #     # Find all numeric patterns (with or without %)
# #     matches = re.findall(r'(\d+\.?\d*)%?', cell_str)
    
# #     valid_numbers = []
# #     for match in matches:
# #         try:
# #             num = float(match)
# #             if 0 < num < 1:  # Handle 0.495 → convert to 49.5
# #                 num *= 100
# #             valid_numbers.append(num)
# #         except ValueError:
# #             continue
    
# #     if not valid_numbers:
# #         return None
    
# #     return min(valid_numbers)


# # def parse_age_based_values(cell_str: str) -> Dict[str, Dict]:
# #     """
# #     Parse cells with age-based patterns like:
# #     - 'Age 0: 27.5% Age 1+: 26%'
# #     - 'Age 0-2: 50% Age 3+: 60%'
# #     - 'AGE0:27.5% AGE1+:26%' (no spaces, typos)
    
# #     Returns dict: {'Age 0': {'payin': 27.5, 'from': 0, 'to': 0}, ...}
# #     """
# #     if not cell_str:
# #         return {}
    
# #     # Normalize: collapse whitespace
# #     cell_str = re.sub(r'\s+', ' ', cell_str.strip())
    
# #     age_values = {}
    
# #     # Pattern 1: Age 0 or Age 0-X (range)
# #     pattern_range = r'age\s*(\d+)\s*[-–]\s*(\d+)\s*[:\-]?\s*([\d\.]+)%?'
# #     for match in re.finditer(pattern_range, cell_str, re.IGNORECASE):
# #         age_from = int(match.group(1))
# #         age_to = int(match.group(2))
# #         val_str = match.group(3)
# #         val = safe_float(val_str)
# #         if val is not None:
# #             label = f"Age {age_from}-{age_to}"
# #             age_values[label] = {"payin": val, "from": age_from, "to": age_to}
    
# #     # Pattern 2: Age X+ (e.g., Age 1+, Age 3+)
# #     pattern_plus = r'age\s*(\d+)\s*\+\s*[:\-]?\s*([\d\.]+)%?'
# #     for match in re.finditer(pattern_plus, cell_str, re.IGNORECASE):
# #         age_num = int(match.group(1))
# #         val_str = match.group(2)
# #         val = safe_float(val_str)
# #         if val is not None:
# #             label = f"Age {age_num}+"
# #             age_values[label] = {"payin": val, "from": age_num, "to": 99}
    
# #     # Pattern 3: Age 0 (single age without + or range)
# #     pattern_single = r'age\s*(\d+)\s*[:\-]?\s*([\d\.]+)%?'
# #     for match in re.finditer(pattern_single, cell_str, re.IGNORECASE):
# #         age_num = int(match.group(1))
# #         val_str = match.group(2)
# #         label_check = f"Age {age_num}"
# #         # Only add if we haven't already captured this as a range or plus
# #         if not any(label_check in k or f"{age_num}+" in k or f"{age_num}-" in k for k in age_values.keys()):
# #             val = safe_float(val_str)
# #             if val is not None:
# #                 age_values[label_check] = {"payin": val, "from": age_num, "to": age_num}
    
# #     return age_values

# # # ===============================================================================
# # # PATTERN DETECTION
# # # ===============================================================================

# # class PatternDetector:
# #     """Detects which pattern the Excel sheet follows."""
    
# #     @staticmethod
# #     def detect_pattern(df: pd.DataFrame) -> str:
# #         """
# #         Detect the pattern type based on sheet structure.
# #         Returns one of: 'april', 'may', 'june', 'july', 'august', 'september', 'standard'
# #         """
# #         # Check for header row location and structure
# #         header_found = False
# #         header_row_idx = None
        
# #         # Search for header in first 15 rows
# #         for i in range(min(15, len(df))):
# #             row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" 
# #                        for j in range(min(df.shape[1], 20))]
            
# #             # Check if this row contains typical headers
# #             if any("CLUSTER" in v or "RTO" in v for v in row_vals):
# #                 header_found = True
# #                 header_row_idx = i
# #                 break
        
# #         if not header_found:
# #             return "standard"  # Fallback to standard pattern
        
# #         # Check row above header for policy types
# #         has_multi_policy = False
# #         if header_row_idx > 0:
# #             prev_row = [str(df.iloc[header_row_idx - 1, j]).upper().strip() 
# #                        if pd.notna(df.iloc[header_row_idx - 1, j]) else "" 
# #                        for j in range(min(df.shape[1], 20))]
            
# #             # Count how many policy type indicators we see
# #             policy_indicators = sum(1 for v in prev_row if "COMP" in v or "SATP" in v or "TP" in v)
# #             if policy_indicators >= 2:
# #                 has_multi_policy = True
        
# #         # Check for segment group row (2 rows above header)
# #         has_segment_groups = False
# #         if header_row_idx >= 2:
# #             segment_row = [str(df.iloc[header_row_idx - 2, j]).upper().strip() 
# #                           if pd.notna(df.iloc[header_row_idx - 2, j]) else "" 
# #                           for j in range(min(df.shape[1], 20))]
            
# #             # Look for segment indicators
# #             segment_indicators = sum(1 for v in segment_row 
# #                                    if len(v) > 2 and v not in ["CD1", "CD2"])
# #             if segment_indicators >= 1:
# #                 has_segment_groups = True
        
# #         # Pattern determination logic
# #         if has_segment_groups and has_multi_policy:
# #             return "july"  # Most complex: segment groups + multiple policies
# #         elif has_multi_policy:
# #             return "may"  # Multiple policy types
# #         elif header_row_idx > 2:
# #             return "june"  # Header lower in sheet
# #         else:
# #             return "april"  # Simplest pattern
    
# #     @staticmethod
# #     def detect_pattern_name(df: pd.DataFrame) -> str:
# #         """Get a descriptive name for the detected pattern."""
# #         pattern = PatternDetector.detect_pattern(df)
# #         pattern_names = {
# #             "april": "April Pattern (Simple With/Without Addon)",
# #             "may": "May Pattern (Multiple Policies)",
# #             "june": "June Pattern (Age-Based Values)",
# #             "july": "July Pattern (Segment Groups + Multi-Policy + Age Bands)",
# #             "august": "August Pattern (Enhanced Multi-Age)",
# #             "september": "September Pattern (Advanced)",
# #             "standard": "Standard Pattern (Basic)"
# #         }
# #         return pattern_names.get(pattern, "Unknown Pattern")

# # # ===============================================================================
# # # PATTERN PROCESSORS
# # # ===============================================================================

# # class AprilPattern:
# #     """Process April pattern: Simple with/without addon pattern."""
    
# #     @staticmethod
# #     def process(df: pd.DataFrame, sheet_name: str) -> List[Dict]:
# #         """Process April pattern sheets."""
# #         records = []
        
# #         try:
# #             print(f"\n{'='*80}")
# #             print(f"Processing Sheet: {sheet_name} (April Pattern)")
# #             print(f"{'='*80}")
            
# #             # Find header row
# #             header_row = None
# #             rto_col = segment_col = make_col = cd1_col = cd2_col = None
            
# #             for i in range(min(10, len(df))):
# #                 for j in range(df.shape[1]):
# #                     cell_val = str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else ""
                    
# #                     if "RTO" in cell_val and "CLUSTER" in cell_val:
# #                         header_row = i
# #                         rto_col = j
# #                     elif "SEGMENT" == cell_val and header_row == i:
# #                         segment_col = j
# #                     elif "MAKE" == cell_val and header_row == i:
# #                         make_col = j
# #                     elif "CD2" == cell_val and header_row == i:
# #                         cd2_col = j
            
# #             if header_row is None:
# #                 print("❌ ERROR: Could not find header row")
# #                 return []
            
# #             print(f"✓ Header found at row {header_row + 1}")
            
# #             # Determine policy type from title row
# #             policy_type = "Comp"
# #             if header_row > 0 and cd2_col is not None:
# #                 title_cell = str(df.iloc[header_row - 1, cd2_col]).upper() if pd.notna(df.iloc[header_row - 1, cd2_col]) else ""
# #                 if "TP" in title_cell:
# #                     policy_type = "TP"
            
# #             # Process data rows
# #             for idx in range(header_row + 1, len(df)):
# #                 row = df.iloc[idx]
                
# #                 rto_cluster = str(row.iloc[rto_col]).strip() if rto_col is not None and pd.notna(row.iloc[rto_col]) else ""
# #                 if not rto_cluster or rto_cluster.lower() in ["", "nan", "none"]:
# #                     continue
                
# #                 segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
# #                 make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
# #                 state = STATE_MAPPING.get(rto_cluster, rto_cluster.upper())
                
# #                 # Process CD2 column
# #                 if cd2_col is not None and cd2_col < len(row):
# #                     cd2_cell = str(row.iloc[cd2_col]).strip() if pd.notna(row.iloc[cd2_col]) else ""
                    
# #                     if cd2_cell:
# #                         # Check for "With Addon" / "Without Addon" pattern
# #                         with_match = re.search(r'with\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
# #                         without_match = re.search(r'without\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
                        
# #                         values_to_process = []
                        
# #                         if with_match:
# #                             val = safe_float(with_match.group(1))
# #                             if val is not None:
# #                                 values_to_process.append((val, "CD2 With Addon"))
                        
# #                         if without_match:
# #                             val = safe_float(without_match.group(1))
# #                             if val is not None:
# #                                 values_to_process.append((val, "CD2 Without Addon"))
                        
# #                         if not with_match and not without_match:
# #                             val = safe_float(cd2_cell)
# #                             if val is not None:
# #                                 values_to_process.append((val, "CD2"))
                        
# #                         # Create records
# #                         for val, addon_type in values_to_process:
# #                             lob = "CV"
# #                             segment_final = "All GVW & PCV 3W, GCV 3W"
# #                             payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_type, val)
                            
# #                             records.append({
# #                                 "State": state,
# #                                 "Location/Cluster": rto_cluster,
# #                                 "Original Segment": segment,
# #                                 "Mapped Segment": segment_final,
# #                                 "LOB": lob,
# #                                 "Policy Type": policy_type,
# #                                 "Payin": f"{val:.2f}%",
# #                                 "Payin Category": get_payin_category(val),
# #                                 "Calculated Payout": f"{payout:.2f}%",
# #                                 "Formula Used": formula,
# #                                 "Rule Explanation": rule_exp,
# #                                 "Remarks": f"{addon_type} | Segment: {segment} | Make: {make}"
# #                             })
            
# #             print(f"✓ Extracted {len(records)} records")
# #             return records
            
# #         except Exception as e:
# #             print(f"❌ ERROR in April pattern processing: {e}")
# #             import traceback
# #             traceback.print_exc()
# #             return []


# # class MayPattern:
# #     """Process May pattern: Multiple policy types with lowest value selection."""
    
# #     @staticmethod
# #     def process(df: pd.DataFrame, sheet_name: str) -> List[Dict]:
# #         """Process May pattern sheets."""
# #         records = []
        
# #         try:
# #             print(f"\n{'='*80}")
# #             print(f"Processing Sheet: {sheet_name} (May Pattern)")
# #             print(f"{'='*80}")
            
# #             # Find header row
# #             header_row = None
# #             policy_row = None
# #             rto_col = segment_col = make_col = None
# #             cd2_cols = []  # List of (column_index, policy_type)
            
# #             # First pass: Find header row
# #             for i in range(min(10, len(df))):
# #                 for j in range(df.shape[1]):
# #                     cell_val = str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else ""
                    
# #                     if "RTO" in cell_val and "CLUSTER" in cell_val:
# #                         header_row = i
# #                         rto_col = j
# #                         break
# #                 if header_row is not None:
# #                     break
            
# #             if header_row is None:
# #                 print("❌ ERROR: Could not find header row")
# #                 return []
            
# #             # Check row above header for policy types
# #             if header_row > 0:
# #                 policy_row = header_row - 1
            
# #             # Second pass: Find all column headers
# #             for j in range(df.shape[1]):
# #                 cell_val = str(df.iloc[header_row, j]).upper().strip() if pd.notna(df.iloc[header_row, j]) else ""
                
# #                 if "SEGMENT" == cell_val:
# #                     segment_col = j
# #                 elif "MAKE" == cell_val:
# #                     make_col = j
# #                 elif "CD2" == cell_val:
# #                     # Get policy type from row above
# #                     policy_type = "Comp"
# #                     if policy_row is not None and pd.notna(df.iloc[policy_row, j]):
# #                         policy_cell = str(df.iloc[policy_row, j]).upper().strip()
# #                         if "SATP" in policy_cell:
# #                             policy_type = "SATP"
# #                         elif "TP" in policy_cell and "SATP" not in policy_cell:
# #                             policy_type = "TP"
# #                         elif "COMP" in policy_cell:
# #                             policy_type = "Comp"
# #                     cd2_cols.append((j, policy_type))
            
# #             print(f"✓ Header found at row {header_row + 1}")
# #             print(f"✓ Found {len(cd2_cols)} CD2 columns with policy types")
            
# #             # Process data rows
# #             for idx in range(header_row + 1, len(df)):
# #                 row = df.iloc[idx]
                
# #                 rto_cluster = str(row.iloc[rto_col]).strip() if rto_col is not None and pd.notna(row.iloc[rto_col]) else ""
# #                 if not rto_cluster or rto_cluster.lower() in ["", "nan", "none"]:
# #                     continue
                
# #                 segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
# #                 make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
# #                 state = STATE_MAPPING.get(rto_cluster, rto_cluster.upper())
                
# #                 # Process each CD2 column
# #                 for cd2_col, policy_type in cd2_cols:
# #                     if cd2_col >= len(row):
# #                         continue
                    
# #                     cell_value = str(row.iloc[cd2_col]).strip() if pd.notna(row.iloc[cd2_col]) else ""
                    
# #                     if cell_value:
# #                         # Extract lowest value from cell (handles formats like "49.5%/44.5%")
# #                         val = extract_lowest_payin(cell_value)
                        
# #                         if val is not None:
# #                             lob = "CV"
# #                             segment_final = "All GVW & PCV 3W, GCV 3W"
# #                             payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_type, val)
                            
# #                             records.append({
# #                                 "State": state,
# #                                 "Location/Cluster": rto_cluster,
# #                                 "Original Segment": segment,
# #                                 "Mapped Segment": segment_final,
# #                                 "LOB": lob,
# #                                 "Policy Type": policy_type,
# #                                 "Payin": f"{val:.2f}%",
# #                                 "Payin Category": get_payin_category(val),
# #                                 "Calculated Payout": f"{payout:.2f}%",
# #                                 "Formula Used": formula,
# #                                 "Rule Explanation": rule_exp,
# #                                 "Remarks": f"Segment: {segment} | Make: {make} | Raw: {cell_value}"
# #                             })
            
# #             print(f"✓ Extracted {len(records)} records")
# #             return records
            
# #         except Exception as e:
# #             print(f"❌ ERROR in May pattern processing: {e}")
# #             import traceback
# #             traceback.print_exc()
# #             return []


# # class JulyPattern:
# #     """Process July pattern: Most complex with segment groups, multiple policies, and age bands."""
    
# #     @staticmethod
# #     def process(df: pd.DataFrame, sheet_name: str) -> List[Dict]:
# #         """Process July pattern sheets."""
# #         records = []
        
# #         try:
# #             print(f"\n{'='*80}")
# #             print(f"Processing Sheet: {sheet_name} (July Pattern - Enhanced)")
# #             print(f"{'='*80}")
            
# #             cluster_col = segment_col = make_col = None
# #             bottom_header_row = mid_header_row = top_header_row = data_start_row = None
            
# #             # Find header rows
# #             for i in range(min(30, len(df))):
# #                 row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" 
# #                            for j in range(df.shape[1])]
                
# #                 if any("CLUSTER" in v or "RTO" in v for v in row_vals):
# #                     bottom_header_row = i
# #                     mid_header_row = i - 1 if i > 0 else None
# #                     top_header_row = i - 2 if i >= 2 else None
                    
# #                     for j, val in enumerate(row_vals):
# #                         # Flexible matching for location column
# #                         if re.search(r'cluster|rto|location|l?cation|loc', val, re.IGNORECASE):
# #                             cluster_col = j
# #                         elif val == "SEGMENT":
# #                             segment_col = j
# #                         elif val == "MAKE":
# #                             make_col = j
                    
# #                     data_start_row = i + 1
# #                     break
            
# #             if cluster_col is None:
# #                 print("❌ Could not find Location/Cluster column")
# #                 return []
            
# #             print(f"✓ Header structure found at row {bottom_header_row + 1}")
            
# #             # Forward-fill merged segment groups (COMP / SATP)
# #             segment_group_map = {}
# #             current_segment = "Unknown"
# #             for j in range(df.shape[1]):
# #                 if top_header_row is not None and pd.notna(df.iloc[top_header_row, j]):
# #                     val = str(df.iloc[top_header_row, j]).strip()
# #                     if val:
# #                         current_segment = val
# #                 segment_group_map[j] = current_segment
            
# #             # Find CD2 columns
# #             cd2_columns = []
# #             for j in range(df.shape[1]):
# #                 bottom_val = str(df.iloc[bottom_header_row, j]).upper().strip() if pd.notna(df.iloc[bottom_header_row, j]) else ""
# #                 if bottom_val != "CD2":
# #                     continue
                
# #                 policy_type = "Comp"
# #                 if mid_header_row is not None and pd.notna(df.iloc[mid_header_row, j]):
# #                     mid_val = str(df.iloc[mid_header_row, j]).upper().strip()
# #                     if "SATP" in mid_val:
# #                         policy_type = "SATP"
                
# #                 segment_group = segment_group_map.get(j, "Unknown")
# #                 cd2_columns.append((j, policy_type, segment_group))
            
# #             print(f"✓ Found {len(cd2_columns)} CD2 columns")
            
# #             # Process data rows
# #             for idx in range(data_start_row, len(df)):
# #                 row = df.iloc[idx]
                
# #                 cluster_raw = row.iloc[cluster_col] if pd.notna(row.iloc[cluster_col]) else ""
# #                 cluster = str(cluster_raw).strip()
# #                 if not cluster or cluster.lower() in ["nan", ""]:
# #                     continue
                
# #                 segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
# #                 make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
# #                 state = STATE_MAPPING.get(cluster, cluster.upper().replace(" ", "_"))
                
# #                 for cd2_col_idx, policy_type, segment_group in cd2_columns:
# #                     cell_value = row.iloc[cd2_col_idx]
# #                     raw_str = str(cell_value) if pd.notna(cell_value) else ""
# #                     raw_str = raw_str.strip()
                    
# #                     if not raw_str:
# #                         continue
                    
# #                     # Skip referrals
# #                     if re.search(r"refer|grid.*refer|above.*rate", raw_str, re.IGNORECASE):
# #                         continue
                    
# #                     # Try enhanced age parsing
# #                     age_dict = parse_age_based_values(raw_str)
                    
# #                     if age_dict:
# #                         # Process each age band
# #                         for age_label, info in age_dict.items():
# #                             payin = info["payin"]
# #                             lob = "CV"
# #                             mapped_segment = "All GVW & PCV 3W, GCV 3W"
# #                             payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_type, payin)
                            
# #                             records.append({
# #                                 "State": state,
# #                                 "Location/Cluster": cluster,
# #                                 "Original Segment": segment,
# #                                 "Mapped Segment": mapped_segment,
# #                                 "Segment Group": segment_group,
# #                                 "Make": make,
# #                                 "Age Label": age_label,
# #                                 "Age From": info["from"],
# #                                 "Age To": info["to"],
# #                                 "LOB": lob,
# #                                 "Policy Type": policy_type,
# #                                 "Payin": f"{payin:.2f}%",
# #                                 "Payin Category": get_payin_category(payin),
# #                                 "Calculated Payout": f"{payout:.2f}%",
# #                                 "Formula Used": formula,
# #                                 "Rule Explanation": rule_exp,
# #                                 "Addon Type": "Age-Based Split",
# #                                 "Remarks": f"Multi-Age Band | Raw: '{raw_str}'"
# #                             })
# #                         continue  # Skip fallback if age bands found
                    
# #                     # Fallback: single value
# #                     val = extract_lowest_payin(raw_str)
# #                     if val is None:
# #                         continue
                    
# #                     addon_type = "Plain"
# #                     if re.search(r'with\s+addon', raw_str, re.IGNORECASE):
# #                         addon_type = "With Addon"
# #                     elif re.search(r'without\s+addon', raw_str, re.IGNORECASE):
# #                         addon_type = "Without Addon"
                    
# #                     lob = "CV"
# #                     mapped_segment = "All GVW & PCV 3W, GCV 3W"
# #                     payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_type, val)
                    
# #                     records.append({
# #                         "State": state,
# #                         "Location/Cluster": cluster,
# #                         "Original Segment": segment,
# #                         "Mapped Segment": mapped_segment,
# #                         "Segment Group": segment_group,
# #                         "Make": make,
# #                         "Age Label": "All Ages",
# #                         "Age From": "",
# #                         "Age To": "",
# #                         "LOB": lob,
# #                         "Policy Type": policy_type,
# #                         "Payin": f"{val:.2f}%",
# #                         "Payin Category": get_payin_category(val),
# #                         "Calculated Payout": f"{payout:.2f}%",
# #                         "Formula Used": formula,
# #                         "Rule Explanation": rule_exp,
# #                         "Addon Type": addon_type,
# #                         "Remarks": f"Single Value | Raw: '{raw_str}'"
# #                     })
            
# #             print(f"✓ Extracted {len(records)} records")
# #             return records
            
# #         except Exception as e:
# #             print(f"❌ ERROR in July pattern processing: {e}")
# #             import traceback
# #             traceback.print_exc()
# #             return []

# # # ===============================================================================
# # # PATTERN DISPATCHER
# # # ===============================================================================

# # class PatternDispatcher:
# #     """Main dispatcher that routes to appropriate pattern processor."""
    
# #     PATTERN_PROCESSORS = {
# #         "april": AprilPattern,
# #         "may": MayPattern,
# #         "june": MayPattern,  # June uses May pattern processor
# #         "july": JulyPattern,
# #         "august": JulyPattern,  # August uses July pattern processor
# #         "september": JulyPattern,  # September uses July pattern processor
# #         "standard": AprilPattern  # Standard falls back to April
# #     }
    
# #     @staticmethod
# #     def process_sheet(df: pd.DataFrame, sheet_name: str) -> List[Dict]:
# #         """
# #         Main entry point for processing any sheet.
# #         Automatically detects pattern and routes to appropriate processor.
# #         """
# #         # Detect pattern
# #         pattern = PatternDetector.detect_pattern(df)
# #         pattern_name = PatternDetector.detect_pattern_name(df)
        
# #         print(f"\n🔍 Pattern Detection: {pattern_name}")
        
# #         # Get appropriate processor
# #         processor_class = PatternDispatcher.PATTERN_PROCESSORS.get(pattern, AprilPattern)
        
# #         # Process the sheet
# #         records = processor_class.process(df, sheet_name)
        
# #         return records

# # # ===============================================================================
# # # API ENDPOINTS
# # # ===============================================================================

# # @app.post("/get_sheets")
# # async def get_sheets(file: UploadFile = File(...)):
# #     contents = await file.read()
# #     xls = pd.ExcelFile(io.BytesIO(contents))
# #     return {"sheets": xls.sheet_names}

# # @app.post("/process")
# # async def process(file: UploadFile = File(...), sheet_name: Optional[str] = Form(None)):
# #     contents = await file.read()
# #     xls = pd.ExcelFile(io.BytesIO(contents))
    
# #     records = []
    
# #     if sheet_name == "all":
# #         # Process all sheets
# #         for s in xls.sheet_names:
# #             df = pd.read_excel(xls, sheet_name=s, header=None)
# #             records += PatternDispatcher.process_sheet(df, s)
# #     elif sheet_name:
# #         # Process specific sheet
# #         df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
# #         records = PatternDispatcher.process_sheet(df, sheet_name)
# #     else:
# #         # If no sheet specified, process the first sheet (process file as is)
# #         first_sheet = xls.sheet_names[0]
# #         df = pd.read_excel(xls, sheet_name=first_sheet, header=None)
# #         records = PatternDispatcher.process_sheet(df, first_sheet)
    
# #     # Generate Excel output
# #     output = io.BytesIO()
# #     pd.DataFrame(records).to_excel(output, index=False, sheet_name='Processed')
# #     output.seek(0)
# #     excel_bytes = output.read()
# #     excel_b64 = base64.b64encode(excel_bytes).decode()
    
# #     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# #     filename = f"CV_Processed_Output_{timestamp}.xlsx"
    
# #     return {
# #         "records": records,
# #         "excel_base64": excel_b64,
# #         "filename": filename
# #     }

# # if __name__ == "__main__":
# #     import uvicorn
# #     uvicorn.run(app, host="0.0.0.0", port=8000)
# #!/usr/bin/env python3
# """
# ================================================================================
# COMPLETE CV UNIFIED BACKEND - ALL PATTERNS MERGED
# ================================================================================
# This FastAPI backend combines ALL working CV pattern processors.

# Patterns Included:
# 1. Probus Pattern (digit_cv.py) - Works for: Probus_Jan25, Digit CV 21-02-2025, 
#    DIGIT PROBUS_Feb25, DIGIT PROBUS_Mar 2025, DIGIT_PROBUS_MAY25

# 2. April Pattern (digit_cv_only_for_one_april.py) - Works for: Digit 13-04-2025

# 3. May Pattern 1 (digit_cv_only_for_may_pattern_1.py) - Works for: 
#    DIGIT - 13-05-2025.xlsx (CV Worksheet only)

# 4. May Pattern 2 (digit_cv_only_for_may_pattern_2.py) - Works for: 
#    DIGIT - 13-05-2025.xlsx (HCV Worksheet), DIGIT - 20-05-2025, DIGIT - 24-05-2025

# 5. June Pattern 1 (digit_cv_only_for_june_pattern_1.py) - Works for: 
#    Digit 19-06-2025, Digit 20-06-2025, Digit 24-06-2025, Digit - JUNE 2025 (HCV)

# 6. June Pattern 2 (digit_cv_only_for_june_pattern_2.py) - Works for: 
#    Digit - JUNE 2025 (CV Worksheet - Pending)

# 7. July Pattern 1 (digit_cv_only_for_july_pattern_1.py) - Works for: 
#    Digit 07-07-2025, DIGIT - 18-07-2025

# 8. July/August Pattern 2 (digit_cv_only_for_july_pattern_2.py / august_pattern_1.py) 
#    - Works for: PROBUS_Aug25 (HCV Worksheet)

# 9. September Pattern 1 (digit_cv_only_for_september_pattern_1.py)

# ================================================================================
# """

# from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Query
# from fastapi.responses import JSONResponse, FileResponse
# from fastapi.middleware.cors import CORSMiddleware
# import pandas as pd
# import io
# import base64
# from typing import Optional, List, Dict
# import re
# from datetime import datetime
# import tempfile
# import os
# import traceback

# # ===============================================================================
# # FASTAPI APP SETUP
# # ===============================================================================

# app = FastAPI(title="CV Complete Unified Processor API")

# # Allow frontend
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ===============================================================================
# # FORMULA DATA AND STATE MAPPING (UNIFIED FROM ALL FILES)
# # ===============================================================================

# FORMULA_DATA = [
#     # CV rules (used by all patterns)
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
# ]

# STATE_MAPPING = {
#     "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
#     "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
#     "Surat": "GUJARAT", "Jaipur": "RAJASTHAN", "Lucknow": "UTTAR PRADESH",
#     "Patna": "BIHAR", "Ranchi": "JHARKHAND", "Bhuvaneshwar": "ODISHA",
#     "Srinagar": "JAMMU AND KASHMIR", "Dehradun": "UTTARAKHAND", "Haridwar": "UTTARAKHAND",
#     "Himachal Pradesh": "HIMACHAL PRADESH", "Andaman": "ANDAMAN AND NICOBAR ISLANDS",
#     "Bangalore": "KARNATAKA", "Jharkhand": "JHARKHAND", "Bihar": "BIHAR",
#     "West Bengal": "WEST BENGAL", "WEST BENGAL": "WEST BENGAL", "North Bengal": "WEST BENGAL",
#     "Orissa": "ODISHA", "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT",
#     "ROM1": "REST OF MAHARASHTRA", "ROM2": "REST OF MAHARASHTRA", 
#     "Good Vizag": "ANDHRA PRADESH", "Good TN": "TAMIL NADU", "Kerala": "KERALA",
#     "Good MP": "MADHYA PRADESH", "Good CG": "CHHATTISGARH",
#     "Good RJ": "RAJASTHAN", "Bad RJ": "RAJASTHAN", "RJ REF": "RAJASTHAN",
#     "Good UP": "UTTAR PRADESH", "Bad UP": "UTTAR PRADESH",
#     "Good UK": "UTTARAKHAND", "Bad UK": "UTTARAKHAND",
#     "Punjab": "PUNJAB", "Jammu": "JAMMU AND KASHMIR", "Assam": "ASSAM",
#     "NE EX ASSAM": "NORTH EAST", "NE excl Assam": "NORTH EAST",
#     "Good NL": "NAGALAND", "GOOD KA": "KARNATAKA", "BAD KA": "KARNATAKA",
#     "HR Ref": "HARYANA", "Dehradun, Haridwar": "UTTARAKHAND"
# }

# # ===============================================================================
# # CORE UTILITY FUNCTIONS (UNIFIED)
# # ===============================================================================

# def safe_float(value):
#     """Safely convert value to float"""
#     if pd.isna(value):
#         return None
#     val_str = str(value).strip().upper()
#     if val_str in ["D", "NA", "", "NAN", "NONE"]:
#         return None
#     try:
#         num = float(val_str.replace('%', '').strip())
#         if 0 < num < 1:
#             num = num * 100
#         return num
#     except:
#         return None


# def get_payin_category(payin: float):
#     """Categorize payin percentage"""
#     if payin <= 20:
#         return "Payin Below 20%"
#     elif payin <= 30:
#         return "Payin 21% to 30%"
#     elif payin <= 50:
#         return "Payin 31% to 50%"
#     else:
#         return "Payin Above 50%"


# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float):
#     """Calculate payout with formula"""
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0, so Payout is 0"
    
#     payin_category = get_payin_category(payin)
#     matching_rule = None
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment:
#             if rule["REMARKS"] == payin_category:
#                 matching_rule = rule
#                 break
    
#     if not matching_rule:
#         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#         payout = round(payin - deduction, 2)
#         return payout, f"-{deduction}%", f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"
    
#     formula = matching_rule["PO"]
#     if formula.startswith("-") and "%" in formula:
#         deduction = float(formula.replace("%", "").replace("-", ""))
#         payout = round(payin - deduction, 2)
#         return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"
#     else:
#         deduction = 2
#         payout = round(payin - deduction, 2)
#         return payout, f"-{deduction}%", f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {payin_category}"


# def extract_lowest_payin(cell_value):
#     """Extract lowest numeric value from string like '15%/10%'"""
#     if pd.isna(cell_value):
#         return None
    
#     cell_str = str(cell_value).strip()
#     if not cell_str or cell_str.upper() in ["D", "NA", "NAN", "NONE", ""]:
#         return None
    
#     matches = re.findall(r'(\d+\.?\d*)%?', cell_str)
#     valid_nums = []
#     for m in matches:
#         try:
#             num = float(m)
#             if 0 < num < 1:
#                 num *= 100
#             valid_nums.append(num)
#         except:
#             continue
    
#     return min(valid_nums) if valid_nums else None


# def parse_age_based_values(cell_str):
#     """
#     Parse age-based values: 'Age 0: 27.5%\nAge 1+: 26%'
#     Returns: {'Age 0': 27.5, 'Age 1+': 26.0}
#     """
#     if not cell_str:
#         return {}
    
#     cell_str = str(cell_str)
#     age_values = {}
    
#     # Pattern for Age 0
#     age0_match = re.search(r'Age\s*0\s*[:\-]?\s*([0-9\.\%/]+)', cell_str, re.IGNORECASE)
#     if age0_match:
#         val = extract_lowest_payin(age0_match.group(1))
#         if val is not None:
#             age_values['Age 0'] = val
    
#     # Pattern for Age 1+
#     age1_match = re.search(r'Age\s*1\s*\+\s*[:\-]?\s*([0-9\.\%/]+)', cell_str, re.IGNORECASE)
#     if age1_match:
#         val = extract_lowest_payin(age1_match.group(1))
#         if val is not None:
#             age_values['Age 1+'] = val
    
#     return age_values

# # ===============================================================================
# # PATTERN 1: PROBUS PATTERN (from digit_cv.py)
# # Works for: Probus_Jan25, Digit CV 21-02-2025, DIGIT PROBUS_Feb25, etc.
# # ===============================================================================

# def process_probus_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
#     """
#     Probus Pattern - Region-based vertical layout
#     - Row 0: Region names
#     - Row 1: Policy types (COMP/TP)
#     - Row 2: CD1/CD2 labels
#     - Row 3+: Data rows with Cluster, Segment, Age, Make columns
#     """
#     records = []
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)

#         # Map each column to its region (from row 0)
#         region_map = {}
#         if len(df) > 0:
#             for col in range(df.shape[1]):
#                 region = str(df.iloc[0, col]).strip() if pd.notna(df.iloc[0, col]) else ""
#                 if region:
#                     region_map[col] = region

#         # Find CD2 columns
#         cd2_columns = []
#         if len(df) > 2:
#             for col in range(df.shape[1]):
#                 cd_label = str(df.iloc[2, col]).upper().strip() if pd.notna(df.iloc[2, col]) else ""
#                 if cd_label == "CD2":
#                     cd2_columns.append(col)

#         # Data starts from row 3
#         for idx in range(3, len(df)):
#             row = df.iloc[idx]
#             cluster = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
#             if not cluster or cluster.lower() in ["", "nan"]:
#                 continue

#             segment = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
#             age_info = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "All"
#             make = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "All"

#             # Process only CD2 columns
#             for cd2_col in cd2_columns:
#                 if cd2_col >= len(row):
#                     continue

#                 region = region_map.get(cd2_col, "UNKNOWN")
#                 state = next((v for k, v in STATE_MAPPING.items() if k.upper() in region.upper()), region.upper())

#                 cell_value = row.iloc[cd2_col]
#                 if pd.isna(cell_value):
#                     continue
#                 cell_str = str(cell_value).strip()

#                 # Determine policy type from column header above
#                 policy_header = str(df.iloc[1, cd2_col]).upper() if len(df) > 1 else ""
#                 policy_type = "Comp" if "COMP" in policy_header else "TP"

#                 payin_list = []

#                 # Case 1: Age conditions → "Age 0-2: 50% Age 3+: 60%"
#                 age_matches = re.finditer(r'Age\s*\d*\-?\d*\s*[:\-]\s*(\d+\.?\d*)%?|Age\s*\d*\+?\s*[:\-]\s*(\d+\.?\d*)%', cell_str, re.IGNORECASE)
#                 for match in age_matches:
#                     for g in match.groups():
#                         if g:
#                             val = safe_float(g)
#                             if val is not None:
#                                 payin_list.append((val, match.group(0).strip()))

#                 # Case 2: X%/Y% → take the SMALLER value
#                 slash_matches = re.findall(r'(\d+\.?\d*)\s*%?\s*/\s*(\d+\.?\d*)%?', cell_str)
#                 for a, b in slash_matches:
#                     v1 = safe_float(a)
#                     v2 = safe_float(b)
#                     if v1 is not None and v2 is not None:
#                         smaller = min(v1, v2)
#                         payin_list.append((smaller, f"{v1:.1f}%/{v2:.1f}% → Smaller: {smaller:.1f}%"))

#                 # Case 3: Single percentage
#                 if not payin_list:
#                     single = safe_float(cell_str)
#                     if single is not None:
#                         payin_list.append((single, ""))

#                 # Fallback: extract all numbers
#                 if not payin_list:
#                     nums = re.findall(r'(\d+\.?\d+)%?', cell_str)
#                     for n in nums:
#                         val = safe_float(n)
#                         if val is not None:
#                             payin_list.append((val, ""))

#                 # Create record for each payin
#                 for payin, remark_text in payin_list:
#                     lob_final = override_lob if override_enabled and override_lob else "CV"
#                     segment_final = override_segment if override_enabled and override_segment else "All GVW & PCV 3W, GCV 3W"
#                     policy_type_final = override_policy_type if override_enabled and override_policy_type else policy_type

#                     payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)

#                     base_remark = f"Cluster: {cluster}"
#                     if segment: base_remark += f" | Segment: {segment}"
#                     if make != "All": base_remark += f" | Make: {make}"
#                     if age_info != "All": base_remark += f" | Age: {age_info}"
#                     if remark_text: base_remark += f" | {remark_text}"

#                     records.append({
#                         "State": state.upper(),
#                         "Location/Cluster": region,
#                         "Original Segment": cluster,
#                         "Mapped Segment": segment_final,
#                         "LOB": lob_final,
#                         "Policy Type": policy_type_final,
#                         "Payin (CD2)": f"{payin:.2f}%",
#                         "Payin Category": get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": formula,
#                         "Rule Explanation": rule_exp,
#                         "Remarks": base_remark.strip()
#                     })

#         return records

#     except Exception as e:
#         print(f"Error in Probus pattern processing: {e}")
#         traceback.print_exc()
#         return []


# # ===============================================================================
# # PATTERN 2: APRIL PATTERN (from digit_cv_only_for_one_april.py)
# # Works for: Digit 13-04-2025
# # ===============================================================================

# def process_april_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
#     """
#     April Pattern - RTO Cluster with 'With Addon' / 'Without Addon'
#     - Headers: RTO Cluster | Segment | Make | CD1 | CD2
#     - Cells contain: "With Addon: 85% Without Addon: 60%" or plain numbers
#     """
#     records = []
    
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         # Find the header row dynamically
#         header_row = None
#         rto_col = None
#         segment_col = None
#         make_col = None
#         cd1_col = None
#         cd2_col = None
        
#         for i in range(min(10, len(df))):
#             for j in range(df.shape[1]):
#                 cell_val = str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else ""
                
#                 if "RTO" in cell_val and "CLUSTER" in cell_val:
#                     header_row = i
#                     rto_col = j
#                 elif "SEGMENT" == cell_val and header_row == i:
#                     segment_col = j
#                 elif "MAKE" == cell_val and header_row == i:
#                     make_col = j
#                 elif "CD1" == cell_val and header_row == i:
#                     cd1_col = j
#                 elif "CD2" == cell_val and header_row == i:
#                     cd2_col = j
        
#         if header_row is None:
#             print("April pattern: Could not find header row with 'RTO Cluster'")
#             return []
        
#         # Check policy type from row above headers
#         policy_type = "Comp"
#         if header_row > 0 and cd1_col is not None:
#             title_cell = str(df.iloc[header_row - 1, cd1_col]).upper() if pd.notna(df.iloc[header_row - 1, cd1_col]) else ""
#             if "TP" in title_cell:
#                 policy_type = "TP"
#             elif "SATP" in title_cell:
#                 policy_type = "SATP"
        
#         # Process data rows
#         for idx in range(header_row + 1, len(df)):
#             row = df.iloc[idx]
            
#             rto_cluster = str(row.iloc[rto_col]).strip() if rto_col is not None and pd.notna(row.iloc[rto_col]) else ""
            
#             if not rto_cluster or rto_cluster.lower() in ["", "nan", "none"]:
#                 continue
            
#             segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
#             make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
#             state = STATE_MAPPING.get(rto_cluster, rto_cluster.upper())
            
#             # Process CD2 column (CD1 processing similar, but focusing on CD2)
#             if cd2_col is not None and cd2_col < len(row):
#                 cd2_cell = str(row.iloc[cd2_col]).strip() if pd.notna(row.iloc[cd2_col]) else ""
                
#                 if cd2_cell:
#                     with_match = re.search(r'with\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
#                     without_match = re.search(r'without\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
                    
#                     values_to_process = []
                    
#                     if with_match:
#                         val = safe_float(with_match.group(1))
#                         if val is not None:
#                             values_to_process.append((val, "CD2 With Addon"))
                    
#                     if without_match:
#                         val = safe_float(without_match.group(1))
#                         if val is not None:
#                             values_to_process.append((val, "CD2 Without Addon"))
                    
#                     if not with_match and not without_match:
#                         val = safe_float(cd2_cell)
#                         if val is not None:
#                             values_to_process.append((val, "CD2"))
                    
#                     for val, remark_prefix in values_to_process:
#                         lob = override_lob if override_enabled and override_lob else "CV"
#                         segment_final = override_segment if override_enabled and override_segment else "All GVW & PCV 3W, GCV 3W"
#                         policy_final = override_policy_type if override_enabled and override_policy_type else policy_type
                        
#                         payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_final, val)
                        
#                         records.append({
#                             "State": state,
#                             "Location/Cluster": rto_cluster,
#                             "Original Segment": segment,
#                             "Mapped Segment": segment_final,
#                             "LOB": lob,
#                             "Policy Type": policy_final,
#                             "Payin (CD2)": f"{val:.2f}%",
#                             "Payin Category": get_payin_category(val),
#                             "Calculated Payout": f"{payout:.2f}%",
#                             "Formula Used": formula,
#                             "Rule Explanation": rule_exp,
#                             "Remarks": f"{remark_prefix} | Segment: {segment} | Make: {make}"
#                         })
        
#         return records
    
#     except Exception as e:
#         print(f"Error in April pattern processing: {e}")
#         traceback.print_exc()
#         return []


# # ===============================================================================
# # PATTERN 3: AUGUST/JULY PATTERN 2 (Age-based with referrals)
# # Works for: PROBUS_Aug25 (HCV), July pattern 2 files
# # ===============================================================================

# def process_august_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type):
#     """
#     August Pattern - Age-based values with segment groups
#     - Multi-header: Top: Segment groups, Mid: Comp/SATP, Bottom: CD2
#     - Supports: Age 0: X%, Age 1+: Y%
#     - Supports: Referral detection
#     """
#     records = []
    
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         cluster_col = segment_col = make_col = age_from_col = age_to_col = None
#         bottom_header_row = mid_header_row = top_header_row = data_start_row = None
        
#         for i in range(min(20, len(df))):
#             row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" for j in range(df.shape[1])]
            
#             if any("CLUSTER" in v for v in row_vals):
#                 bottom_header_row = i
#                 mid_header_row = i - 1 if i > 0 else None
#                 top_header_row = i - 2 if i >= 2 else None
                
#                 for j, val in enumerate(row_vals):
#                     if "CLUSTER" in val:
#                         cluster_col = j
#                     elif val == "SEGMENT":
#                         segment_col = j
#                     elif val == "MAKE":
#                         make_col = j
#                     elif "AGE FROM" in val:
#                         age_from_col = j
#                     elif "AGE TO" in val:
#                         age_to_col = j
                
#                 data_start_row = i + 1
#                 break
        
#         if bottom_header_row is None:
#             print("August pattern: Could not find header row with 'Cluster'")
#             return []
        
#         # Forward-fill merged segment groups
#         segment_group_map = {}
#         current_segment = None
#         for j in range(df.shape[1]):
#             cell_val = ""
#             if top_header_row is not None and pd.notna(df.iloc[top_header_row, j]):
#                 cell_val = str(df.iloc[top_header_row, j]).strip()
#             if cell_val:
#                 current_segment = cell_val
#             segment_group_map[j] = current_segment or "Unknown Segment"
        
#         # Find CD2 columns
#         cd2_columns = []
#         for j in range(df.shape[1]):
#             bottom_val = str(df.iloc[bottom_header_row, j]).upper().strip() if pd.notna(df.iloc[bottom_header_row, j]) else ""
#             if bottom_val != "CD2":
#                 continue
            
#             policy_type = "Comp"
#             if mid_header_row is not None and pd.notna(df.iloc[mid_header_row, j]):
#                 mid_val = str(df.iloc[mid_header_row, j]).upper().strip()
#                 if "SATP" in mid_val:
#                     policy_type = "SATP"
#                 elif "COMP" in mid_val:
#                     policy_type = "Comp"
            
#             segment_group = segment_group_map.get(j, "Unknown Segment")
#             cd2_columns.append((j, policy_type, segment_group))
        
#         for idx in range(data_start_row, len(df)):
#             row = df.iloc[idx]
            
#             cluster = str(row.iloc[cluster_col]).strip() if cluster_col is not None and pd.notna(row.iloc[cluster_col]) else ""
#             if not cluster or cluster.lower() in ["", "nan"]:
#                 continue
            
#             segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
#             make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
            
#             state = STATE_MAPPING.get(cluster, cluster.upper())
            
#             for cd2_col_idx, policy_type, segment_group in cd2_columns:
#                 cell_value = row.iloc[cd2_col_idx]
#                 raw_str = str(cell_value) if pd.notna(cell_value) else ""
#                 raw_str = raw_str.strip()
                
#                 if not raw_str:
#                     continue
                
#                 # Check for referral
#                 if re.search(r"refer|grids?.*refer|above.*rates?", raw_str, re.IGNORECASE):
#                     # Try to find previous value
#                     found = False
#                     for prev_idx in range(idx - 1, data_start_row - 1, -1):
#                         val = extract_lowest_payin(df.iloc[prev_idx, cd2_col_idx])
#                         if val is not None:
#                             final_payin = val
#                             addon_type = "Referred"
#                             age_label = "Referred Age"
#                             found = True
#                             break
#                     if not found:
#                         continue
#                 else:
#                     # Try age-based parsing
#                     age_dict = parse_age_based_values(raw_str)
                    
#                     if age_dict:
#                         for age_key, payin in age_dict.items():
#                             lob = override_lob if override_enabled and override_lob else "CV"
#                             mapped_segment = override_segment if override_enabled and override_segment else "All GVW & PCV 3W, GCV 3W"
#                             policy_final = override_policy_type if override_enabled and override_policy_type else policy_type
                            
#                             payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_final, payin)
                            
#                             age_from_out = "0" if age_key == "Age 0" else "1"
#                             age_to_out = "0" if age_key == "Age 0" else "99"
                            
#                             records.append({
#                                 "State": state,
#                                 "Location/Cluster": cluster,
#                                 "Original Segment": segment,
#                                 "Mapped Segment": mapped_segment,
#                                 "LOB": lob,
#                                 "Policy Type": policy_final,
#                                 "Payin (CD2)": f"{payin:.2f}%",
#                                 "Payin Category": get_payin_category(payin),
#                                 "Calculated Payout": f"{payout:.2f}%",
#                                 "Formula Used": formula,
#                                 "Rule Explanation": rule_exp,
#                                 "Remarks": f"Age-Based | {age_key} | {segment_group} | Make: {make}"
#                             })
#                         # Skip fallback if age-based was found
#                         continue
                
#                 # Fallback: normal single value
#                 val = extract_lowest_payin(raw_str)
#                 if val is None:
#                     continue
                
#                 with_match = re.search(r'with\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
#                 without_match = re.search(r'without\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
#                 addon_type = ("With Addon (Lowest)" if with_match else 
#                              "Without Addon" if without_match else 
#                              "Plain/Multiple (Lowest)")
                
#                 final_payin = val
                
#                 lob = override_lob if override_enabled and override_lob else "CV"
#                 mapped_segment = override_segment if override_enabled and override_segment else "All GVW & PCV 3W, GCV 3W"
#                 policy_final = override_policy_type if override_enabled and override_policy_type else policy_type
                
#                 payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_final, final_payin)
                
#                 records.append({
#                     "State": state,
#                     "Location/Cluster": cluster,
#                     "Original Segment": segment,
#                     "Mapped Segment": mapped_segment,
#                     "LOB": lob,
#                     "Policy Type": policy_final,
#                     "Payin (CD2)": f"{final_payin:.2f}%",
#                     "Payin Category": get_payin_category(final_payin),
#                     "Calculated Payout": f"{payout:.2f}%",
#                     "Formula Used": formula,
#                     "Rule Explanation": rule_exp,
#                     "Remarks": f"{addon_type} | {segment_group} | Make: {make}"
#                 })
        
#         return records
        
#     except Exception as e:
#         print(f"Error in August pattern processing: {e}")
#         traceback.print_exc()
#         return []


# # ===============================================================================
# # PATTERN DETECTION LOGIC
# # ===============================================================================

# def detect_cv_pattern(df, sheet_name):
#     """
#     Detect which CV pattern the sheet follows
#     Returns pattern name as string
#     """
#     # Sample first 20 rows for detection
#     text = " | ".join(df.head(20).astype(str).stack().str.upper().tolist())
#     sheet_lower = sheet_name.lower()
    
#     # Pattern 1: April - "With Addon" / "Without Addon"
#     if "RTO" in text and "CLUSTER" in text and ("WITH ADDON" in text or "WITHOUT ADDON" in text):
#         print(f"Detected: April Pattern (RTO with Addon)")
#         return "april"
    
#     # Pattern 2: August/July - Age-based with segment groups
#     if ("AGE 0:" in text or "AGE 1+" in text or "AGE 1 +" in text) and ("NON-DUMPER" in text or "TIPPER" in text or "NON DUMPER" in text):
#         print(f"Detected: August Pattern (Age-based with segments)")
#         return "august"
    
#     # Pattern 3: Probus - Region-based vertical (most common)
#     # This is the fallback for most CV/HCV sheets
#     if "CLUSTER" in text and "CD2" in text:
#         print(f"Detected: Probus Pattern (Region-based vertical)")
#         return "probus"
    
#     print(f"No specific pattern detected, defaulting to Probus")
#     return "probus"


# # ===============================================================================
# # FILE STORAGE
# # ===============================================================================

# uploaded_files = {}

# # ===============================================================================
# # API ENDPOINTS
# # ===============================================================================

# @app.get("/")
# async def root():
#     """API information"""
#     return {
#         "name": "CV Complete Unified Processor API",
#         "version": "3.0.0",
#         "description": "All CV patterns merged into one backend",
#         "patterns": [
#             "Probus (digit_cv.py)",
#             "April (digit_cv_only_for_one_april.py)",
#             "August/July (digit_cv_only_for_august_pattern_1.py)",
#         ],
#         "endpoints": {
#             "POST /upload": "Upload Excel file",
#             "POST /process": "Process with auto-detection",
#             "POST /export": "Export to Excel",
#             "GET /health": "Health check"
#         }
#     }


# @app.get("/health")
# async def health_check():
#     """Health check"""
#     return {
#         "status": "healthy",
#         "timestamp": datetime.now().isoformat(),
#         "version": "3.0.0"
#     }


# @app.post("/upload")
# async def upload_file(file: UploadFile = File(...)):
#     """Upload Excel file and detect worksheets"""
#     try:
#         content = await file.read()
#         xls = pd.ExcelFile(io.BytesIO(content))
#         sheets = xls.sheet_names
        
#         sheet_info = []
#         for sheet in sheets:
#             # Determine sheet type
#             sheet_lower = sheet.lower()
#             sheet_type = "cv" if any(kw in sheet_lower for kw in ["cv", "hcv", "commercial", "probus"]) else "unknown"
            
#             # Read preview
#             try:
#                 df_preview = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None, nrows=5)
#                 preview = {
#                     "columns": [str(c) for c in df_preview.columns[:5]],
#                     "sample_data": [[str(df_preview.iloc[i, j]) if pd.notna(df_preview.iloc[i, j]) else "" 
#                                     for j in range(min(5, df_preview.shape[1]))] 
#                                    for i in range(min(3, len(df_preview)))]
#                 }
#             except:
#                 preview = {"columns": [], "sample_data": []}
            
#             sheet_info.append({
#                 "name": sheet,
#                 "type": sheet_type,
#                 "type_display": "Commercial Vehicle" if sheet_type == "cv" else "Unknown",
#                 "icon": "🚛" if sheet_type == "cv" else "📄",
#                 "preview": preview
#             })
        
#         # Store file
#         file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
#         uploaded_files[file_id] = {
#             "content": content,
#             "filename": file.filename,
#             "sheets": sheets,
#             "sheet_info": sheet_info
#         }
        
#         auto_selected = len(sheets) == 1
#         auto_selected_sheet = sheets[0] if auto_selected else None
        
#         return {
#             "file_id": file_id,
#             "filename": file.filename,
#             "total_sheets": len(sheets),
#             "sheet_info": sheet_info,
#             "auto_selected": auto_selected,
#             "auto_selected_sheet": auto_selected_sheet,
#             "message": f"File uploaded. Found {len(sheets)} worksheet(s)." + 
#                       (f" Auto-selected '{auto_selected_sheet}'." if auto_selected else " Select a worksheet to process.")
#         }
        
#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Error uploading file: {str(e)}")


# @app.post("/process")
# async def process_sheet(
#     file_id: str,
#     sheet_name: str,
#     override_enabled: bool = False,
#     override_lob: Optional[str] = None,
#     override_segment: Optional[str] = None,
#     override_policy_type: Optional[str] = None
# ):
#     """Process worksheet with automatic pattern detection"""
#     try:
#         if file_id not in uploaded_files:
#             raise HTTPException(status_code=404, detail="File not found. Please upload again.")
        
#         file_data = uploaded_files[file_id]
#         content = file_data["content"]
        
#         if sheet_name not in file_data["sheets"]:
#             raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found")
        
#         # Read sheet
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         # Detect pattern
#         pattern = detect_cv_pattern(df, sheet_name)
        
#         print(f"Processing: {sheet_name}, Pattern: {pattern}")
        
#         # Route to appropriate processor
#         if pattern == "april":
#             records = process_april_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type)
#             processor_name = "CV - April Pattern (RTO with Addon)"
#         elif pattern == "august":
#             records = process_august_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type)
#             processor_name = "CV - August Pattern (Age-based)"
#         else:  # probus
#             records = process_probus_pattern(content, sheet_name, override_enabled, override_lob, override_segment, override_policy_type)
#             processor_name = "CV - Probus Pattern (Vertical)"
        
#         if not records:
#             return {
#                 "success": False,
#                 "message": "No records extracted. Check sheet structure or use override settings.",
#                 "records": [],
#                 "count": 0,
#                 "processor": processor_name
#             }
        
#         # Calculate summary
#         states = {}
#         lobs = {}
#         policies = {}
#         payins = []
#         payouts = []
        
#         for record in records:
#             state = record.get("State", "Unknown")
#             states[state] = states.get(state, 0) + 1
            
#             lob = record.get("LOB", "Unknown")
#             lobs[lob] = lobs.get(lob, 0) + 1
            
#             policy = record.get("Policy Type", "Unknown")
#             policies[policy] = policies.get(policy, 0) + 1
            
#             try:
#                 payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
#                 payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
#                 if payin > 0:
#                     payins.append(payin)
#                     payouts.append(payout)
#             except:
#                 pass
        
#         avg_payin = sum(payins) / len(payins) if payins else 0
#         avg_payout = sum(payouts) / len(payouts) if payouts else 0
        
#         summary = {
#             "total_records": len(records),
#             "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
#             "lobs": lobs,
#             "policies": policies,
#             "average_payin": round(avg_payin, 2),
#             "average_payout": round(avg_payout, 2),
#             "processor": processor_name,
#             "pattern": pattern,
#             "sheet_name": sheet_name
#         }
        
#         return {
#             "success": True,
#             "message": f"Successfully processed {len(records)} records using {processor_name}",
#             "records": records,
#             "count": len(records),
#             "summary": summary
#         }
        
#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Error processing: {str(e)}")


# @app.post("/export")
# async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
#     """Export processed records to Excel"""
#     try:
#         if not records:
#             raise HTTPException(status_code=400, detail="No records to export")
        
#         df = pd.DataFrame(records)
        
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         filename = f"CV_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
#         temp_dir = tempfile.gettempdir()
#         output_path = os.path.join(temp_dir, filename)
        
#         df.to_excel(output_path, index=False, sheet_name='Processed')
        
#         return FileResponse(
#             path=output_path,
#             filename=filename,
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error exporting: {str(e)}")


# if __name__ == "__main__":
#     import uvicorn
#     print("=" * 80)
#     print("CV COMPLETE UNIFIED PROCESSOR - ALL PATTERNS MERGED")
#     print("=" * 80)
#     print("Included Patterns:")
#     print("  1. Probus Pattern (Region-based vertical)")
#     print("  2. April Pattern (RTO with Addon)")
#     print("  3. August Pattern (Age-based with segments)")
#     print("=" * 80)
#     print("Starting server on http://0.0.0.0:8000")
#     print("=" * 80)
#     uvicorn.run(app, host="0.0.0.0", port=8000)
# from fastapi import FastAPI, File, UploadFile, Form, HTTPException
# from fastapi.responses import JSONResponse
# from fastapi.middleware.cors import CORSMiddleware
# import pandas as pd
# import io
# import base64
# from typing import Optional, List, Dict
# import re

# app = FastAPI(title="Insurance Policy Processor API - Unified CV Edition")

# # Allow frontend (localhost:5500 or any)
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ------------------- FORMULA DATA -------------------
# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
#     # TW SAOD + COMP rules
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     # TW TP rules
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
#     # PVT CAR rules
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     # CV rules
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     # BUS rules
#     {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
#     {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
#     # TAXI rules
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     # MISD rules
#     {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
# ]

# # Merged STATE_MAPPING from all patterns
# STATE_MAPPING = {
#     "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
#     "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
#     "Bihar": "BIHAR", "Jharkhand": "JHARKHAND", "Patna": "BIHAR", "Ranchi": "JHARKHAND",
#     "ROM2": "DELHI", "Punjab": "PUNJAB", "NE excl Assam": "NORTH EAST", "Good RJ": "RAJASTHAN",
#     "Bad RJ": "RAJASTHAN", "RJ REF": "RAJASTHAN", "Andaman": "ANDAMAN AND NICOBAR",
#     "ROM1": "REST OF MAHARASHTRA",
#     "Good CG": "CHHATTISGARH", "Bad CG": "CHHATTISGARH",
# }

# def safe_float(value):
#     """
#     Safely convert a cell value (string, float, etc.) to a percentage float.
#     Handles common cases: "20%", 0.20 → 20.0, "NA", empty, "D", negative values → None
#     """
#     if pd.isna(value):
#         return None
    
#     val_str = str(value).strip().upper()
    
#     # Explicit skip values
#     if val_str in ["D", "NA", "NAN", "NONE", "", "DECLINE", "-"]:
#         return None
    
#     try:
#         # Remove % if present
#         cleaned = val_str.replace('%', '').replace(' ', '')
        
#         num = float(cleaned)
        
#         # Convert decimal form (0.20 → 20.0)
#         if 0 < num <= 1:
#             return round(num * 100, 2)
        
#         # Already in percentage form (20 → 20.0, 27.5 → 27.5)
#         if num > 1:
#             return round(num, 2)
        
#         # Zero is allowed (though rare)
#         if num == 0:
#             return 0.0
            
#         return None  # negative or invalid
        
#     except (ValueError, TypeError):
#         return None
# def extract_state(cluster_name: str) -> str:
#     """
#     Extract state from cluster / RTO name using the mapping dictionary.
#     Falls back to 'REST OF INDIA' or 'UNKNOWN'.
#     """
#     if pd.isna(cluster_name) or not cluster_name:
#         return "UNKNOWN"
    
#     cluster_upper = str(cluster_name).upper().strip()
    
#     for key, mapped_state in STATE_MAPPING.items():
#         if key.upper() in cluster_upper:
#             return mapped_state
    
#     # Try to take first meaningful part before any delimiter
#     parts = re.split(r'[-_/ ,]+', cluster_upper)
#     for part in parts:
#         if part in STATE_MAPPING:
#             return STATE_MAPPING[part]
    
#     return "REST OF INDIA"   # or "UNKNOWN" — your choice
# # ------------------- HELPER FUNCTIONS FROM ALL PATTERNS -------------------
# def get_payin_category(payin: float):
#     if payin <= 20:
#         return "Payin Below 20%"
#     elif payin <= 30:
#         return "Payin 21% to 30%"
#     elif payin <= 50:
#         return "Payin 31% to 50%"
#     else:
#         return "Payin Above 50%"

# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float):
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0, so Payout is 0"
    
#     payin_category = get_payin_category(payin)
#     matching_rule = None
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment:
#             if rule["REMARKS"] == payin_category:
#                 matching_rule = rule
#                 break
    
#     if not matching_rule:
#         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#         payout = round(payin - deduction, 2)
#         return payout, f"-{deduction}%", f"Default deduction applied for {payin_category}"
    
#     formula = matching_rule["PO"]
#     deduction = float(formula.replace("%", "").replace("-", ""))
#     payout = round(payin - deduction, 2)
#     return payout, formula, f"Matched rule: {payin_category}"

# def extract_lowest_payin(cell_value):
#     if pd.isna(cell_value):
#         return None
    
#     cell_str = str(cell_value).strip()
#     if not cell_str or cell_str.upper() in ["D", "NA", "NAN", "NONE", "", "D"]:
#         return None
    
#     matches = re.findall(r'(\d+\.?\d*)%?', cell_str)
#     valid_nums = []
#     for m in matches:
#         try:
#             num = float(m)
#             if 0 < num < 1:
#                 num *= 100
#             valid_nums.append(num)
#         except:
#             continue
    
#     return min(valid_nums) if valid_nums else None

# def parse_age_based_values(cell_str):
#     if not cell_str:
#         return {}
    
#     cell_str = str(cell_str)
#     age_values = {}
    
#     patterns = [
#         r'Age\s*(\d+)-(\d+)\s*[:\-]?\s*([0-9\.\%/]+)',
#         r'Age\s*(\d+)\s*\+\s*[:\-]?\s*([0-9\.\%/]+)',
#         r'Age\s*(\d+)\s*[:\-]?\s*([0-9\.\%/]+)'  # fallback: Age 0: xx%
#     ]

#     for pattern in patterns:
#         matches = re.finditer(pattern, cell_str, re.IGNORECASE)
#         for match in matches:
#             if len(match.groups()) == 3:  # Age 0-2
#                 age_from, age_to, val_str = match.groups()
#                 label = f"Age {age_from}-{age_to}"
#                 age_from_out = age_from
#                 age_to_out = age_to
#             elif len(match.groups()) == 2:  # Age 3+ or Age 0:
#                 if '+' in match.group(0):
#                     age_from = match.group(1)
#                     label = f"Age {age_from}+"
#                     age_from_out = age_from
#                     age_to_out = "99"
#                     val_str = match.group(2)
#                 else:
#                     age_num = match.group(1)
#                     label = f"Age {age_num}"
#                     age_from_out = age_num
#                     age_to_out = age_num
#                     val_str = match.group(2)
#             else:
#                 continue

#             val = extract_lowest_payin(val_str)
#             if val is not None:
#                 age_values[label] = {
#                     "value": val,
#                     "from": age_from_out,
#                     "to": age_to_out
#                 }
#     return age_values

# # ------------------- PATTERN DETECTION -------------------
# def detect_cv_pattern(df: pd.DataFrame, sheet_name: str = "") -> str:
#     df_str = df.to_string().upper()
#     sheet_lower = sheet_name.lower()
    
#     if "WITH ADDON" in df_str or "WITHOUT ADDON" in df_str:
#         return "addon"
    
#     if "AGE 0" in df_str or "AGE 1+" in df_str or "AGE FROM" in df_str or "AGE TO" in df_str:
#         return "age_split"
    
#     if "STATE" in df_str and "REGION" in df_str:
#         return "main"
    
#     # More specific for other patterns
#     if "RTO CLUSTER" in df_str and "SEGMENT" in df_str and "MAKE" in df_str:
#         return "addon"  # for may_pattern1 and april
    
#     return "main"  # fallback

# # ------------------- PROCESS FUNCTIONS FROM EACH PATTERN -------------------
# # Main pattern from digit_cv.py
# def process_cv_main(content, sheet, override_enabled, override_lob, override_segment, override_policy_type):
#     records = []
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
#         header_row = None
#         for i in range(10):
#             row_vals = [str(val).upper() for val in df.iloc[i] if pd.notna(val)]
#             if "REGION" in row_vals or "CLUSTER" in row_vals:
#                 header_row = i
#                 break
#         if header_row is None:
#             header_row = 0
        
#         df.columns = df.iloc[header_row]
#         df = df.iloc[header_row + 1:].reset_index(drop=True)
#         df.columns = [str(col).strip().upper() if pd.notna(col) else f"UNNAMED_{i}" for i, col in enumerate(df.columns)]
        
#         state_col = next((col for col in df.columns if "STATE" in col), None)
#         region_col = next((col for col in df.columns if "REGION" in col), None)
#         cluster_col = next((col for col in df.columns if "CLUSTER" in col), None)
#         lob_col = next((col for col in df.columns if "LOB" in col), None)
#         policy_type_col = next((col for col in df.columns if "POLICY TYPE" in col), None)
#         payin_col = next((col for col in df.columns if "PAYIN" in col or "CD2" in col), None)
#         segment_col = next((col for col in df.columns if "SEGMENT" in col), None)
#         remarks_col = next((col for col in df.columns if "REMARKS" in col), None)
        
#         for _, row in df.iterrows():
#             region = str(row[region_col]).strip() if region_col and pd.notna(row[region_col]) else ""
#             cluster = str(row[cluster_col]).strip() if cluster_col and pd.notna(row[cluster_col]) else ""
#             if not cluster:
#                 continue
#             state = str(row[state_col]).upper() if state_col and pd.notna(row[state_col]) else extract_state(cluster)
#             lob = str(row[lob_col]).upper() if lob_col and pd.notna(row[lob_col]) else "CV"
#             policy_type = str(row[policy_type_col]).upper() if policy_type_col and pd.notna(row[policy_type_col]) else "TP"
#             payin_str = str(row[payin_col]).strip() if payin_col and pd.notna(row[payin_col]) else ""
#             payin = safe_float(payin_str)
#             if payin is None or payin <= 0:
#                 continue
#             segment = str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else "All GVW & PCV 3W, GCV 3W"
#             lob_final = override_lob if override_enabled and override_lob else lob
#             segment_final = override_segment if override_enabled and override_segment else segment
#             policy_type_final = override_policy_type if override_policy_type else policy_type
#             payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)
#             remark_text = str(row[remarks_col]).strip() if remarks_col and pd.notna(row[remarks_col]) else ""
#             base_remark = f"Segment: {segment} | Cluster: {cluster}"
#             if remark_text:
#                 base_remark += f" | {remark_text}"
            
#             records.append({
#                 "State": state,
#                 "Location/Cluster": region,
#                 "Original Segment": cluster,
#                 "Mapped Segment": segment_final,
#                 "LOB": lob_final,
#                 "Policy Type": policy_type_final,
#                 "Payin (CD2)": f"{payin:.2f}%",
#                 "Payin Category": get_payin_category(payin),
#                 "Calculated Payout": f"{payout:.2f}%",
#                 "Formula Used": formula,
#                 "Rule Explanation": rule_exp,
#                 "Remarks": base_remark.strip()
#             })
        
#         return records
    
#     except Exception as e:
#         print(f"Error in main CV processing: {e}")
#         return []

# # Addon pattern from digit_cv_only_for_may_pattern_1.py
# def process_cv_addon(content, sheet, override_enabled, override_lob, override_segment, override_policy_type):
#     records = []
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
        
#         header_row = None
#         policy_row = None
#         rto_col = None
#         segment_col = None
#         make_col = None
#         cd1_cols = []  # List of tuples: (column_index, policy_type)
#         cd2_cols = []  # List of tuples: (column_index, policy_type)
        
#         # First pass: Find header row
#         for i in range(min(10, len(df))):
#             for j in range(df.shape[1]):
#                 cell_val = str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else ""
                
#                 if "RTO" in cell_val and "CLUSTER" in cell_val:
#                     header_row = i
#                     rto_col = j
#                 elif "SEGMENT" == cell_val and header_row == i:
#                     segment_col = j
#                 elif "MAKE" == cell_val and header_row == i:
#                     make_col = j
#                 elif "CD1" == cell_val and header_row == i:
#                     cd1_cols.append((j, "Comp"))
#                 elif "CD2" == cell_val and header_row == i:
#                     cd2_cols.append((j, "Comp"))
        
#         if header_row is None:
#             print("❌ ERROR: Could not find header row with 'RTO Cluster'")
#             return []
        
#         print(f"\n📍 Header found at Row {header_row + 1}")
#         print(f"   RTO Cluster: Column {rto_col} ({chr(65+rto_col)})")
#         print(f"   Segment: Column {segment_col} ({chr(65+segment_col)})" if segment_col else "   Segment: Not found")
#         print(f"   Make: Column {make_col} ({chr(65+make_col)})" if make_col else "   Make: Not found")
#         print(f"   CD1 Columns: {cd1_cols}")
#         print(f"   CD2 Columns: {cd2_cols}")
        
#         # Find policy type row (above header)
#         policy_row = header_row - 1 if header_row > 0 else None
#         if policy_row is not None:
#             policy_types = df.iloc[policy_row]
#             for col_idx, policy_type in enumerate(policy_types):
#                 policy_str = str(policy_type).strip().upper() if pd.notna(policy_type) else ""
#                 if policy_str in ["COMP", "TP", "SAOD"]:
#                     # Assign policy type to CD1 and CD2 columns
#                     cd1_cols.append((col_idx, policy_str))
#                     cd2_cols.append((col_idx, policy_str))
        
#         # Data start row
#         data_start = header_row + 1
        
#         # Process rows
#         rows_processed = 0
#         current_state = ""
#         for i in range(data_start, len(df)):
#             row = df.iloc[i]
#             rto_cluster = str(row[rto_col]).strip() if pd.notna(row[rto_col]) else ""
#             if not rto_cluster or rto_cluster.upper() in ["TOTAL", "GRAND TOTAL", ""]:
#                 continue
#             rows_processed += 1
            
#             print(f"Processing row {i+1}: {rto_cluster}")
            
#             state = extract_state(rto_cluster)
#             if state == "UNKNOWN" and current_state:
#                 state = current_state
#             else:
#                 current_state = state
            
#             segment = str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else ""
#             make = str(row[make_col]).strip() if make_col and pd.notna(row[make_col]) else ""
            
#             for col_idx, policy_type in cd2_cols:
#                 cd2_cell = row[col_idx]
#                 raw_cell_str = str(cd2_cell).strip() if pd.notna(cd2_cell) else ""
#                 if not raw_cell_str:
#                     continue
                
#                 # Parse with/without addon
#                 with_match = re.search(r'with\s+addon\s*[:\s]*([0-9\.\%\/]+)', raw_cell_str, re.IGNORECASE)
#                 without_match = re.search(r'without\s+addon\s*[:\s]*([0-9\.\%\/]+)', raw_cell_str, re.IGNORECASE)
                
#                 processed = False
                
#                 if with_match:
#                     val = extract_lowest_payin(with_match.group(1))
#                     if val is not None:
#                         lob = "CV"
#                         segment_final = "All GVW & PCV 3W, GCV 3W"
#                         payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_type, val)
                        
#                         print(f"    ✓ With Addon (lowest): {val}% → Payout: {payout}% (Formula: {formula})")
                        
#                         records.append({
#                             "State": state,
#                             "Location/Cluster": rto_cluster,
#                             "Original Segment": segment,
#                             "Mapped Segment": segment_final,
#                             "LOB": lob,
#                             "Policy Type": policy_type,
#                             "Payin": f"{val:.2f}%",
#                             "Payin Category": get_payin_category(val),
#                             "Calculated Payout": f"{payout:.2f}%",
#                             "Formula Used": formula,
#                             "Rule Explanation": rule_exp,
#                             "Remarks": f"CD2 With Addon (Lowest) | Raw: '{raw_cell_str}' | Segment: {segment} | Make: {make}"
#                         })
#                     processed = True
                
#                 if without_match:
#                     val = extract_lowest_payin(without_match.group(1))
#                     if val is not None:
#                         lob = "CV"
#                         segment_final = "All GVW & PCV 3W, GCV 3W"
#                         payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_type, val)
                        
#                         print(f"    ✓ Without Addon (lowest): {val}% → Payout: {payout}% (Formula: {formula})")
                        
#                         records.append({
#                             "State": state,
#                             "Location/Cluster": rto_cluster,
#                             "Original Segment": segment,
#                             "Mapped Segment": segment_final,
#                             "LOB": lob,
#                             "Policy Type": policy_type,
#                             "Payin": f"{val:.2f}%",
#                             "Payin Category": get_payin_category(val),
#                             "Calculated Payout": f"{payout:.2f}%",
#                             "Formula Used": formula,
#                             "Rule Explanation": rule_exp,
#                             "Remarks": f"CD2 Without Addon (Lowest) | Raw: '{raw_cell_str}' | Segment: {segment} | Make: {make}"
#                         })
#                     processed = True
                
#                 if not processed:
#                     val = extract_lowest_payin(raw_cell_str)
#                     if val is not None:
#                         lob = "CV"
#                         segment_final = "All GVW & PCV 3W, GCV 3W"
#                         payout, formula, rule_exp = calculate_payout_with_formula(lob, segment_final, policy_type, val)
                        
#                         print(f"    ✓ Multiple/Plain value (lowest): {val}% → Payout: {payout}% (Formula: {formula})")
                        
#                         records.append({
#                             "State": state,
#                             "Location/Cluster": rto_cluster,
#                             "Original Segment": segment,
#                             "Mapped Segment": segment_final,
#                             "LOB": lob,
#                             "Policy Type": policy_type,
#                             "Payin": f"{val:.2f}%",
#                             "Payin Category": get_payin_category(val),
#                             "Calculated Payout": f"{payout:.2f}%",
#                             "Formula Used": formula,
#                             "Rule Explanation": rule_exp,
#                             "Remarks": f"CD2 Multiple values (Lowest) | Raw: '{raw_cell_str}' | Segment: {segment} | Make: {make}"
#                         })
            
#         return records
#     except Exception as e:
#         print(f"Error in addon CV processing: {e}")
#         return []

# # Age split pattern from digit_cv_only_for_june_pattern_1.py (enhanced with multi-age band)
# def process_cv_age_split(content, sheet, override_enabled, override_lob, override_segment, override_policy_type):
#     records = []
#     try:
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
        
#         cluster_col = segment_col = make_col = age_from_col = age_to_col = None
#         bottom_header_row = mid_header_row = top_header_row = data_start_row = None
        
#         for i in range(min(20, len(df))):
#             row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" for j in range(df.shape[1])]
            
#             if any("CLUSTER" in v for v in row_vals):
#                 bottom_header_row = i
#                 mid_header_row = i - 1 if i > 0 else None
#                 top_header_row = i - 2 if i >= 2 else i - 3 if i >= 3 else None
                
#                 for j, val in enumerate(row_vals):
#                     if "CLUSTER" in val:
#                         cluster_col = j
#                     elif val == "SEGMENT":
#                         segment_col = j
#                     elif val == "MAKE":
#                         make_col = j
#                     elif "AGE FROM" in val:
#                         age_from_col = j
#                     elif "AGE TO" in val:
#                         age_to_col = j
        
#         if bottom_header_row is None:
#             return []
        
#         if mid_header_row is None:
#             mid_header_row = bottom_header_row - 1
        
#         if top_header_row is None:
#             top_header_row = mid_header_row - 1
        
#         data_start_row = bottom_header_row + 1
        
#         segment_group = ""
#         current_segment_group = ""
#         current_cluster = ""
#         current_segment = ""
#         current_make = ""
        
#         tp_cols = []
#         saod_cols = []
        
#         for j in range(df.shape[1]):
#             bottom_val = str(df.iloc[bottom_header_row, j]).upper().strip()
#             mid_val = str(df.iloc[mid_header_row, j]).upper().strip() if mid_header_row is not None else ""
#             top_val = str(df.iloc[top_header_row, j]).upper().strip() if top_header_row is not None else ""
            
#             if "TP" in bottom_val or "TP" in mid_val or "TP" in top_val:
#                 tp_cols.append(j)
#             elif "SAOD" in bottom_val or "SAOD" in mid_val or "SAOD" in top_val or "COMP" in bottom_val:
#                 saod_cols.append(j)
        
#         policy_cols = {
#             "TP": tp_cols,
#             "SAOD": saod_cols
#         }
        
#         for i in range(data_start_row, len(df)):
#             row = df.iloc[i]
            
#             cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
#             if cluster and cluster.upper() not in ["TOTAL", "GRAND TOTAL", ""]:
#                 current_cluster = cluster
#                 current_segment_group = str(row[top_header_row, cluster_col]).strip() if top_header_row else ""
#             else:
#                 cluster = current_cluster
            
#             segment = str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else ""
#             if segment:
#                 current_segment = segment
            
#             make = str(row[make_col]).strip() if make_col and pd.notna(row[make_col]) else ""
#             if make:
#                 current_make = make
            
#             age_from = str(row[age_from_col]).strip() if age_from_col and pd.notna(row[age_from_col]) else ""
#             age_to = str(row[age_to_col]).strip() if age_to_col and pd.notna(row[age_to_col]) else ""
            
#             if not cluster:
#                 continue
            
#             state = extract_state(cluster)
            
#             for policy_type, cols in policy_cols.items():
#                 for col in cols:
#                     cell_value = str(row[col]).strip() if pd.notna(row[col]) else ""
#                     if not cell_value:
#                         continue
                    
#                     raw_str = cell_value
#                     age_values = parse_age_based_values(raw_str)
                    
#                     if age_values:
#                         for label, data in age_values.items():
#                             val = data["value"]
#                             age_from_out = data["from"]
#                             age_to_out = data["to"]
#                             lob = "CV"
#                             mapped_segment = "All GVW & PCV 3W, GCV 3W"
#                             payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_type, val)
                            
#                             records.append({
#                                 "State": state,
#                                 "Location/Cluster": cluster,
#                                 "Original Segment": segment,
#                                 "Mapped Segment": mapped_segment,
#                                 "Segment Group": segment_group,
#                                 "Make": make,
#                                 "Age Label": label,
#                                 "Age From": age_from_out,
#                                 "Age To": age_to_out,
#                                 "LOB": lob,
#                                 "Policy Type": policy_type,
#                                 "Payin": f"{val:.2f}%",
#                                 "Payin Category": get_payin_category(val),
#                                 "Calculated Payout": f"{payout:.2f}%",
#                                 "Formula Used": formula,
#                                 "Rule Explanation": rule_exp,
#                                 "Addon Type": "Age-Based Split",
#                                 "Remarks": f"Multi-Age Band | Raw: '{raw_str}'"
#                             })
#                         continue  # Skip fallback if age bands found
                    
#                     # Fallback: single value
#                     val = extract_lowest_payin(raw_str)
#                     if val is None:
#                         continue
                    
#                     addon_type = "Plain"
#                     if re.search(r'with\s+addon', raw_str, re.IGNORECASE):
#                         addon_type = "With Addon"
#                     elif re.search(r'without\s+addon', raw_str, re.IGNORECASE):
#                         addon_type = "Without Addon"
                    
#                     lob = "CV"
#                     mapped_segment = "All GVW & PCV 3W, GCV 3W"
#                     payout, formula, rule_exp = calculate_payout_with_formula(lob, mapped_segment, policy_type, val)
                    
#                     records.append({
#                         "State": state,
#                         "Location/Cluster": cluster,
#                         "Original Segment": segment,
#                         "Mapped Segment": mapped_segment,
#                         "Segment Group": segment_group,
#                         "Make": make,
#                         "Age Label": "All Ages",
#                         "Age From": "",
#                         "Age To": "",
#                         "LOB": lob,
#                         "Policy Type": policy_type,
#                         "Payin": f"{val:.2f}%",
#                         "Payin Category": get_payin_category(val),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": formula,
#                         "Rule Explanation": rule_exp,
#                         "Addon Type": addon_type,
#                         "Remarks": f"Single Value | Raw: '{raw_str}'"
#                     })
        
#         return records
    
#     except Exception as e:
#         print(f"Error in age split CV processing: {e}")
#         return []

# # ------------------- MAIN PROCESS ENDPOINT -------------------
# @app.post("/process")
# async def process_file(
#     company_name: str = Form(...),
#     policy_file: UploadFile = File(...),
#     sheet_name: Optional[str] = Form(None),
#     override_enabled: str = Form("false"),
#     override_lob: Optional[str] = Form(None),
#     override_segment: Optional[str] = Form(None),
#     override_policy_type: Optional[str] = Form(None),
# ):
#     try:
#         content = await policy_file.read()
#         xls = pd.ExcelFile(io.BytesIO(content))
#         if sheet_name and sheet_name in xls.sheet_names:
#             sheets_to_process = [sheet_name]
#         else:
#             sheets_to_process = xls.sheet_names
#         all_records = []
#         for sheet in sheets_to_process:
#             print(f"Processing sheet: {sheet}")
#             sheet_lower = sheet.lower()
#             if "cv" in sheet_lower:
#                 df = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
#                 pattern = detect_cv_pattern(df, sheet)
#                 if pattern == "main":
#                     records = process_cv_main(content, sheet, override_enabled, override_lob, override_segment, override_policy_type)
#                 elif pattern == "addon":
#                     records = process_cv_addon(content, sheet, override_enabled, override_lob, override_segment, override_policy_type)
#                 elif pattern == "age_split":
#                     records = process_cv_age_split(content, sheet, override_enabled, override_lob, override_segment, override_policy_type)
#                 else:
#                     records = process_cv_main(content, sheet, override_enabled, override_lob, override_segment, override_policy_type)
#             all_records.extend(records)
#             print(f"Sheet '{sheet}' produced {len(records)} records")
#         if not all_records:
#             raise HTTPException(status_code=400, detail="No valid data found in any sheet")
#         result_df = pd.DataFrame(all_records)
#         payins = [float(r["Payin (CD2)"].replace('%', '')) for r in all_records if float(r["Payin (CD2)"].replace('%', '')) > 0]
#         avg_payin = round(sum(payins) / len(payins), 2) if payins else 0
#         formula_summary = {}
#         for r in all_records: 
#             f = r["Formula Used"]
#             formula_summary[f] = formula_summary.get(f, 0) + 1
#         excel_buffer = io.BytesIO()
#         with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
#             result_df.to_excel(writer, index=False, sheet_name='Processed')
#         excel_buffer.seek(0)
#         excel_b64 = base64.b64encode(excel_buffer.read()).decode()
#         csv_buffer = io.StringIO()
#         result_df.to_csv(csv_buffer, index=False)
#         csv_b64 = base64.b64encode(csv_buffer.getvalue().encode()).decode()
#         json_str = result_df.to_json(orient="records")
#         return {
#             "metrics": {
#                 "company_name": company_name,
#                 "total_records": len(all_records),
#                 "avg_payin": f"{avg_payin:.2f}",
#                 "unique_segments": len(result_df["Mapped Segment"].unique()),
#                 "formula_summary": formula_summary,
#                 "sheets_processed": len(sheets_to_process)
#             },
#             "calculated_data": all_records,
#             "excel_data": excel_b64,
#             "csv_data": csv_b64,
#             "json_data": json_str
#         }
#     except Exception as e:
#         print(f"Error: {str(e)}")
#         raise HTTPException(status_code=500, detail=str(e))

# if __name__ == "__main__":
#     import uvicorn
#     uvicorn.run(app, host="127.0.0.1", port=8000)

# from fastapi import FastAPI, File, UploadFile, HTTPException
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import FileResponse
# import pandas as pd
# import io
# import os
# import re
# from typing import List, Dict, Optional, Tuple
# from datetime import datetime
# import traceback
# import tempfile
# from abc import ABC, abstractmethod

# app = FastAPI(title="DIGIT CV Unified Processor API")

# # Enable CORS
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ===============================================================================
# # FORMULA DATA AND STATE MAPPING
# # ===============================================================================

# FORMULA_DATA = [
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
# ]

# STATE_MAPPING = {
#     "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
#     "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
#     "Bihar": "BIHAR", "Jharkhand": "JHARKHAND", "Patna": "BIHAR", "Ranchi": "JHARKHAND",
#     "ROM2": "DELHI", "Punjab": "PUNJAB", "NE excl Assam": "NORTH EAST", "Good RJ": "RAJASTHAN",
#     "Bad RJ": "RAJASTHAN", "RJ REF": "RAJASTHAN", "Andaman": "ANDAMAN AND NICOBAR",
#     "ROM1": "REST OF MAHARASHTRA", "Surat": "GUJARAT", "Jaipur": "RAJASTHAN",
#     "West Bengal": "WEST BENGAL", "North Bengal": "WEST BENGAL", "Orissa": "ODISHA",
#     "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT", "Good Vizag": "ANDHRA PRADESH",
#     "Good TN": "TAMIL NADU", "Kerala": "KERALA", "Good MP": "MADHYA PRADESH",
#     "Good CG": "CHHATTISGARH", "Good UP": "UTTAR PRADESH", "Bad UP": "UTTAR PRADESH",
#     "Good UK": "UTTARAKHAND", "Bad UK": "UTTARAKHAND", "Jammu": "JAMMU AND KASHMIR",
#     "Assam": "ASSAM", "NE EX ASSAM": "NORTH EAST", "HR Ref": "HARYANA"
# }

# uploaded_files = {}

# # ===============================================================================
# # CORE CALCULATION FUNCTIONS
# # ===============================================================================

# def get_payin_category(payin: float) -> str:
#     if payin <= 20: return "Payin Below 20%"
#     elif payin <= 30: return "Payin 21% to 30%"
#     elif payin <= 50: return "Payin 31% to 50%"
#     else: return "Payin Above 50%"

# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0"
    
#     payin_category = get_payin_category(payin)
#     matching_rule = None
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment:
#             if rule["REMARKS"] == payin_category:
#                 matching_rule = rule
#                 break
    
#     if not matching_rule:
#         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#         payout = round(payin - deduction, 2)
#         return payout, f"-{deduction}%", f"Default: {payin_category}"
    
#     formula = matching_rule["PO"]
#     deduction = float(formula.replace("%", "").replace("-", ""))
#     payout = round(payin - deduction, 2)
#     return payout, formula, f"Matched: {payin_category}"

# def extract_lowest_payin(cell_value) -> Optional[float]:
#     """Extract lowest numeric value from strings like '15%/10%', '30%', etc."""
#     if pd.isna(cell_value):
#         return None
    
#     cell_str = str(cell_value).strip()
#     if not cell_str or cell_str.upper() in ["D", "NA", "NAN", "NONE", "", "DECLINE"]:
#         return None
    
#     matches = re.findall(r'(\d+\.?\d*)%?', cell_str)
#     valid_nums = []
#     for m in matches:
#         try:
#             num = float(m)
#             if 0 < num < 1:
#                 num *= 100
#             valid_nums.append(num)
#         except:
#             continue
    
#     return min(valid_nums) if valid_nums else None

# def safe_float(value) -> Optional[float]:
#     if pd.isna(value):
#         return None
#     val_str = str(value).strip().upper()
#     if val_str in ["D", "NA", "", "NAN", "NONE", "DECLINE"]:
#         return None
#     try:
#         num = float(val_str.replace('%', '').strip())
#         if 0 < num < 1:
#             num = num * 100
#         return num
#     except:
#         return None

# # ===============================================================================
# # BASE PROCESSOR CLASS
# # ===============================================================================

# class CVProcessor(ABC):
#     """Base class for all CV pattern processors"""
    
#     @abstractmethod
#     def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
#         """Detect if this processor can handle the given sheet"""
#         pass
    
#     @abstractmethod
#     def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
#         """Process the sheet and return records"""
#         pass
    
#     @abstractmethod
#     def get_pattern_name(self) -> str:
#         """Return the pattern name for display"""
#         pass

# # ===============================================================================
# # PATTERN 1: BASE/PROBUS PATTERN (Jan-May Pattern)
# # ===============================================================================

# class ProbusPatternProcessor(CVProcessor):
#     """Handles standard Probus pattern with region columns and CD2"""
    
#     def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
#         # Check for region mapping in row 0
#         if df.shape[0] < 3:
#             return False
        
#         # Look for regions like "Delhi", "Mumbai", etc in top rows
#         top_row = df.iloc[0].astype(str).str.upper()
#         has_regions = any(state in ' '.join(top_row.tolist()) for state in ["DELHI", "MUMBAI", "GOOD", "ROM"])
        
#         # Check for CD2 in row 2
#         if df.shape[0] > 2:
#             row2 = df.iloc[2].astype(str).str.upper()
#             has_cd2 = "CD2" in ' '.join(row2.tolist())
#             return has_regions and has_cd2
        
#         return False
    
#     def get_pattern_name(self) -> str:
#         return "Probus Base Pattern (Jan-May)"
    
#     def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
#         records = []
        
#         try:
#             # Map regions from row 0
#             region_map = {}
#             for col in range(df.shape[1]):
#                 val = str(df.iloc[0, col]).strip() if pd.notna(df.iloc[0, col]) else ""
#                 if val:
#                     region_map[col] = val
            
#             # Find CD2 columns from row 2
#             cd2_columns = []
#             for col in range(df.shape[1]):
#                 if pd.notna(df.iloc[2, col]) and "CD2" in str(df.iloc[2, col]).upper():
#                     cd2_columns.append(col)
            
#             # Process data from row 3 onwards
#             for idx in range(3, len(df)):
#                 row = df.iloc[idx]
#                 cluster = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
#                 if not cluster:
#                     continue
                
#                 segment = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
#                 age_info = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "All"
#                 make = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "All"
                
#                 for cd2_col in cd2_columns:
#                     region = region_map.get(cd2_col, "UNKNOWN")
#                     state = STATE_MAPPING.get(region, region.upper())
                    
#                     cell_value = row.iloc[cd2_col]
#                     if pd.isna(cell_value):
#                         continue
                    
#                     # Determine policy type from row 1
#                     policy_header = str(df.iloc[1, cd2_col]).upper() if len(df) > 1 else ""
#                     policy_type = "Comp" if "COMP" in policy_header else "TP"
                    
#                     # Extract payin
#                     payin = safe_float(cell_value)
#                     if payin is None:
#                         continue
                    
#                     payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, payin)
                    
#                     records.append({
#                         "State": state,
#                         "Location/Cluster": region,
#                         "Original Segment": cluster,
#                         "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
#                         "Segment": segment,
#                         "Make": make,
#                         "Age": age_info,
#                         "LOB": "CV",
#                         "Policy Type": policy_type,
#                         "Payin (CD2)": f"{payin:.2f}%",
#                         "Payin Category": get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": formula,
#                         "Rule Explanation": rule_exp
#                     })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in Probus pattern: {e}")
#             traceback.print_exc()
#             return []

# # ===============================================================================
# # PATTERN 2: APRIL PATTERN
# # ===============================================================================

# class AprilPatternProcessor(CVProcessor):
#     """April 2025 specific pattern"""
    
#     def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
#         # Similar to Probus but with specific April characteristics
#         return False  # Will implement specific detection
    
#     def get_pattern_name(self) -> str:
#         return "April 2025 Pattern"
    
#     def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
#         # Same as Probus for now
#         return ProbusPatternProcessor().process(df, sheet_name)

# # ===============================================================================
# # PATTERN 3: MAY PATTERN 1 (CV Worksheet)
# # ===============================================================================

# class MayPattern1Processor(CVProcessor):
#     """May Pattern 1 - CV Worksheet with age-based values"""
    
#     def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
#         # Check for age-based patterns in cells
#         sample_str = df.head(20).to_string().upper()
#         return "AGE 0" in sample_str and "AGE 1+" in sample_str
    
#     def get_pattern_name(self) -> str:
#         return "May Pattern 1 (Age-Based)"
    
#     def parse_age_based_values(self, cell_str: str) -> Dict[str, float]:
#         """Parse 'Age 0: 27.5% Age 1+: 26%' format"""
#         if not cell_str:
#             return {}
        
#         age_values = {}
        
#         # Age 0
#         age0_match = re.search(r'Age\s*0\s*[:\-]?\s*([0-9\.\%/]+)', str(cell_str), re.IGNORECASE)
#         if age0_match:
#             val = extract_lowest_payin(age0_match.group(1))
#             if val is not None:
#                 age_values['Age 0'] = val
        
#         # Age 1+
#         age1_match = re.search(r'Age\s*1\s*\+\s*[:\-]?\s*([0-9\.\%/]+)', str(cell_str), re.IGNORECASE)
#         if age1_match:
#             val = extract_lowest_payin(age1_match.group(1))
#             if val is not None:
#                 age_values['Age 1+'] = val
        
#         return age_values
    
#     def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
#         records = []
        
#         try:
#             # Find header row
#             header_row = None
#             for i in range(min(10, len(df))):
#                 row_str = ' '.join(df.iloc[i].astype(str).str.upper())
#                 if "CLUSTER" in row_str and "CD2" in row_str:
#                     header_row = i
#                     break
            
#             if header_row is None:
#                 return []
            
#             # Find CD2 columns
#             cd2_cols = []
#             for j in range(df.shape[1]):
#                 if pd.notna(df.iloc[header_row, j]) and "CD2" in str(df.iloc[header_row, j]).upper():
#                     cd2_cols.append(j)
            
#             # Process data
#             for idx in range(header_row + 1, len(df)):
#                 row = df.iloc[idx]
#                 cluster = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
#                 if not cluster:
#                     continue
                
#                 state = STATE_MAPPING.get(cluster, cluster.upper())
                
#                 for cd2_col in cd2_cols:
#                     cell_value = row.iloc[cd2_col]
#                     if pd.isna(cell_value):
#                         continue
                    
#                     cell_str = str(cell_value)
#                     age_values = self.parse_age_based_values(cell_str)
                    
#                     for age_label, payin in age_values.items():
#                         payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", "Comp", payin)
                        
#                         records.append({
#                             "State": state,
#                             "Location/Cluster": cluster,
#                             "Original Segment": f"CV - {age_label}",
#                             "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
#                             "Age": age_label,
#                             "LOB": "CV",
#                             "Policy Type": "Comp",
#                             "Payin (CD2)": f"{payin:.2f}%",
#                             "Payin Category": get_payin_category(payin),
#                             "Calculated Payout": f"{payout:.2f}%",
#                             "Formula Used": formula,
#                             "Rule Explanation": rule_exp
#                         })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in May Pattern 1: {e}")
#             traceback.print_exc()
#             return []

# # ===============================================================================
# # PATTERN 4: MAY PATTERN 2 (HCV - Multi-Header)
# # ===============================================================================

# class MayPattern2Processor(CVProcessor):
#     """May Pattern 2 - HCV with segment groups and merged headers"""
    
#     def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
#         # Check for multi-level headers
#         if df.shape[0] < 5:
#             return False
        
#         # Look for segment groups in top rows
#         top_str = ' '.join(df.iloc[:3].to_string().split()).upper()
#         return "NON-DUMPER" in top_str or "TIPPER" in top_str or "SEGMENT" in top_str
    
#     def get_pattern_name(self) -> str:
#         return "May Pattern 2 (HCV Multi-Header)"
    
#     def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
#         records = []
        
#         try:
#             # Find headers
#             bottom_header_row = mid_header_row = top_header_row = None
            
#             for i in range(min(20, len(df))):
#                 row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" for j in range(df.shape[1])]
                
#                 if any("CLUSTER" in v for v in row_vals):
#                     bottom_header_row = i
#                     mid_header_row = i - 1 if i > 0 else None
#                     top_header_row = i - 2 if i >= 2 else None
#                     break
            
#             if bottom_header_row is None:
#                 return []
            
#             # Build segment group mapping (forward fill for merged cells)
#             segment_group_map = {}
#             current_segment = None
            
#             if top_header_row is not None:
#                 for j in range(df.shape[1]):
#                     cell_val = str(df.iloc[top_header_row, j]).strip() if pd.notna(df.iloc[top_header_row, j]) else ""
#                     if cell_val:
#                         current_segment = cell_val
#                     segment_group_map[j] = current_segment or "Unknown Segment"
            
#             # Find CD2 columns
#             cd2_columns = []
#             cluster_col = segment_col = make_col = age_from_col = age_to_col = None
            
#             for j in range(df.shape[1]):
#                 bottom_val = str(df.iloc[bottom_header_row, j]).upper().strip() if pd.notna(df.iloc[bottom_header_row, j]) else ""
                
#                 if "CLUSTER" in bottom_val:
#                     cluster_col = j
#                 elif bottom_val == "SEGMENT":
#                     segment_col = j
#                 elif bottom_val == "MAKE":
#                     make_col = j
#                 elif "AGE FROM" in bottom_val:
#                     age_from_col = j
#                 elif "AGE TO" in bottom_val:
#                     age_to_col = j
#                 elif bottom_val == "CD2":
#                     # Determine policy type
#                     policy_type = "Comp"
#                     if mid_header_row is not None and pd.notna(df.iloc[mid_header_row, j]):
#                         mid_val = str(df.iloc[mid_header_row, j]).upper().strip()
#                         if "SATP" in mid_val:
#                             policy_type = "SATP"
                    
#                     segment_group = segment_group_map.get(j, "Unknown Segment")
#                     cd2_columns.append((j, policy_type, segment_group))
            
#             # Process data rows
#             data_start_row = bottom_header_row + 1
            
#             for idx in range(data_start_row, len(df)):
#                 row = df.iloc[idx]
                
#                 cluster = str(row.iloc[cluster_col]).strip() if cluster_col is not None and pd.notna(row.iloc[cluster_col]) else ""
#                 if not cluster:
#                     continue
                
#                 segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
#                 make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
#                 age_from = row.iloc[age_from_col] if age_from_col is not None else ""
#                 age_to = row.iloc[age_to_col] if age_to_col is not None else ""
                
#                 state = STATE_MAPPING.get(cluster, cluster.upper())
                
#                 for cd2_col_idx, policy_type, segment_group in cd2_columns:
#                     cell_value = row.iloc[cd2_col_idx]
#                     if pd.isna(cell_value):
#                         continue
                    
#                     raw_str = str(cell_value).strip()
                    
#                     # Check for referral notes
#                     if re.search(r"refer|grids?.*to.*be.*refer|above.*rates?", raw_str, re.IGNORECASE):
#                         # Search previous rows for value
#                         found = False
#                         for prev_idx in range(idx - 1, data_start_row - 1, -1):
#                             prev_cell = df.iloc[prev_idx, cd2_col_idx]
#                             val = extract_lowest_payin(prev_cell)
#                             if val is not None:
#                                 payin = val
#                                 addon_type = "Referred from above"
#                                 found = True
#                                 break
#                         if not found:
#                             continue
#                     else:
#                         # Extract with/without addon
#                         with_match = re.search(r'with\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
#                         without_match = re.search(r'without\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
                        
#                         if with_match:
#                             payin = extract_lowest_payin(with_match.group(1))
#                             addon_type = "With Addon"
#                         elif without_match:
#                             payin = extract_lowest_payin(without_match.group(1))
#                             addon_type = "Without Addon"
#                         else:
#                             payin = extract_lowest_payin(cell_value)
#                             addon_type = "Plain"
                        
#                         if payin is None:
#                             continue
                    
#                     payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, payin)
                    
#                     records.append({
#                         "State": state,
#                         "Location/Cluster": cluster,
#                         "Original Segment": segment,
#                         "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
#                         "Segment Group": segment_group,
#                         "Make": make,
#                         "Age From": age_from,
#                         "Age To": age_to,
#                         "LOB": "CV",
#                         "Policy Type": policy_type,
#                         "Payin (CD2)": f"{payin:.2f}%",
#                         "Payin Category": get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": formula,
#                         "Rule Explanation": rule_exp,
#                         "Addon Type": addon_type
#                     })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in May Pattern 2: {e}")
#             traceback.print_exc()
#             return []

# # ===============================================================================
# # PATTERN DETECTOR AND DISPATCHER
# # ===============================================================================

# class CVPatternDetector:
#     """Detects which CV pattern to use based on sheet structure"""
    
#     # All available processors in priority order
#     PROCESSORS = [
#         MayPattern2Processor(),
#         MayPattern1Processor(),
#         ProbusPatternProcessor(),
#         AprilPatternProcessor(),
#     ]
    
#     @staticmethod
#     def detect_pattern(df: pd.DataFrame, sheet_name: str) -> CVProcessor:
#         """Detect and return the appropriate processor"""
        
#         for processor in CVPatternDetector.PROCESSORS:
#             if processor.detect(df, sheet_name):
#                 print(f"✓ Detected: {processor.get_pattern_name()}")
#                 return processor
        
#         # Default to Probus pattern
#         print(f"⚠ No specific pattern detected, using Probus Base Pattern")
#         return ProbusPatternProcessor()
    
#     @staticmethod
#     def process_sheet(df: pd.DataFrame, sheet_name: str) -> Tuple[List[Dict], str]:
#         """Detect pattern and process sheet"""
#         processor = CVPatternDetector.detect_pattern(df, sheet_name)
#         records = processor.process(df, sheet_name)
#         return records, processor.get_pattern_name()

# # ===============================================================================
# # API ENDPOINTS
# # ===============================================================================

# @app.get("/")
# async def root():
#     return {"message": "DIGIT CV Unified Processor API", "version": "2.0"}

# @app.post("/upload")
# async def upload_file(file: UploadFile = File(...)):
#     """Upload an Excel file and return available worksheets"""
#     try:
#         if not file.filename.endswith(('.xlsx', '.xls')):
#             raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
#         content = await file.read()
#         xls = pd.ExcelFile(io.BytesIO(content))
#         sheets = xls.sheet_names
        
#         file_id = datetime.now().strftime("%Y%m%d_%H%M%S")
#         uploaded_files[file_id] = {
#             "content": content,
#             "filename": file.filename,
#             "sheets": sheets
#         }
        
#         return {
#             "file_id": file_id,
#             "filename": file.filename,
#             "sheets": sheets,
#             "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
#         }
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

# @app.post("/process")
# async def process_sheet(
#     file_id: str,
#     sheet_name: str
# ):
#     """Process a specific worksheet with automatic pattern detection"""
#     try:
#         if file_id not in uploaded_files:
#             raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
#         file_data = uploaded_files[file_id]
#         content = file_data["content"]
        
#         if sheet_name not in file_data["sheets"]:
#             raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
#         # Load sheet
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         # Detect pattern and process
#         records, pattern_name = CVPatternDetector.process_sheet(df, sheet_name)
        
#         if not records:
#             return {
#                 "success": False,
#                 "message": "No records extracted. Please check the sheet structure.",
#                 "records": [],
#                 "count": 0,
#                 "pattern": pattern_name
#             }
        
#         # Calculate statistics
#         states = {}
#         policies = {}
#         payins = []
#         payouts = []
        
#         for record in records:
#             state = record.get("State", "Unknown")
#             states[state] = states.get(state, 0) + 1
            
#             policy = record.get("Policy Type", "Unknown")
#             policies[policy] = policies.get(policy, 0) + 1
            
#             try:
#                 payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
#                 payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
#                 if payin > 0:
#                     payins.append(payin)
#                     payouts.append(payout)
#             except:
#                 pass
        
#         avg_payin = sum(payins) / len(payins) if payins else 0
#         avg_payout = sum(payouts) / len(payouts) if payouts else 0
        
#         summary = {
#             "total_records": len(records),
#             "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
#             "policies": policies,
#             "average_payin": round(avg_payin, 2),
#             "average_payout": round(avg_payout, 2),
#             "pattern": pattern_name
#         }
        
#         return {
#             "success": True,
#             "message": f"Successfully processed {len(records)} records using {pattern_name}",
#             "records": records,
#             "count": len(records),
#             "summary": summary
#         }
        
#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")

# @app.post("/export")
# async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
#     """Export processed records to Excel file"""
#     try:
#         if not records:
#             raise HTTPException(status_code=400, detail="No records to export")
        
#         df = pd.DataFrame(records)
        
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         filename = f"CV_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
#         temp_dir = tempfile.gettempdir()
#         output_path = os.path.join(temp_dir, filename)
        
#         df.to_excel(output_path, index=False, sheet_name='Processed')
        
#         return FileResponse(
#             path=output_path,
#             filename=filename,
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")

# if __name__ == "__main__":
#     import uvicorn
#     uvicorn.run(app, host="0.0.0.0", port=8000)


from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import io
import os
import re
from typing import List, Dict, Optional, Tuple
from datetime import datetime
import traceback
import tempfile
from abc import ABC, abstractmethod

app = FastAPI(title="DIGIT CV Unified Processor API - Complete Edition")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://digit-excel-cv.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===============================================================================
# FORMULA DATA AND STATE MAPPING
# ===============================================================================

FORMULA_DATA = [
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
]

STATE_MAPPING = {
    "DELHI": "DELHI", "Mumbai": "MAHARASHTRA", "Pune": "MAHARASHTRA", "Goa": "GOA",
    "Kolkata": "WEST BENGAL", "Hyderabad": "TELANGANA", "Ahmedabad": "GUJARAT",
    "Bihar": "BIHAR", "Jharkhand": "JHARKHAND", "Patna": "BIHAR", "Ranchi": "JHARKHAND",
    "ROM2": "DELHI", "Punjab": "PUNJAB", "NE excl Assam": "NORTH EAST", "Good RJ": "RAJASTHAN",
    "Bad RJ": "RAJASTHAN", "RJ REF": "RAJASTHAN", "Andaman": "ANDAMAN AND NICOBAR",
    "ROM1": "REST OF MAHARASHTRA", "Surat": "GUJARAT", "Jaipur": "RAJASTHAN",
    "West Bengal": "WEST BENGAL", "North Bengal": "WEST BENGAL", "Orissa": "ODISHA",
    "Good GJ": "GUJARAT", "Bad GJ": "GUJARAT", "Good Vizag": "ANDHRA PRADESH",
    "Good TN": "TAMIL NADU", "Kerala": "KERALA", "Good MP": "MADHYA PRADESH",
    "Good CG": "CHHATTISGARH", "Good UP": "UTTAR PRADESH", "Bad UP": "UTTAR PRADESH",
    "Good UK": "UTTARAKHAND", "Bad UK": "UTTARAKHAND", "Jammu": "JAMMU AND KASHMIR",
    "Assam": "ASSAM", "NE EX ASSAM": "NORTH EAST", "HR Ref": "HARYANA",
    "Himachal Pradesh": "HIMACHAL PRADESH", "Bangalore": "KARNATAKA",
    "Bhuvaneshwar": "ODISHA", "Srinagar": "JAMMU AND KASHMIR", "Dehradun": "UTTARAKHAND",
    "Haridwar": "UTTARAKHAND", "Lucknow": "UTTAR PRADESH"
}

uploaded_files = {}

# ===============================================================================
# CORE CALCULATION FUNCTIONS
# ===============================================================================

def get_payin_category(payin: float) -> str:
    """Categorize payin into predefined ranges"""
    if payin <= 20: return "Payin Below 20%"
    elif payin <= 30: return "Payin 21% to 30%"
    elif payin <= 50: return "Payin 31% to 50%"
    else: return "Payin Above 50%"

def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
    """Calculate payout using formula rules"""
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    
    payin_category = get_payin_category(payin)
    matching_rule = None
    
    for rule in FORMULA_DATA:
        if rule["LOB"] == lob and rule["SEGMENT"] == segment:
            if rule["REMARKS"] == payin_category:
                matching_rule = rule
                break
    
    if not matching_rule:
        deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
        payout = round(payin - deduction, 2)
        return payout, f"-{deduction}%", f"Default: {payin_category}"
    
    formula = matching_rule["PO"]
    deduction = float(formula.replace("%", "").replace("-", ""))
    payout = round(payin - deduction, 2)
    return payout, formula, f"Matched: {payin_category}"

def extract_lowest_payin(cell_value) -> Optional[float]:
    """Extract lowest numeric value from strings like '15%/10%', '30%', etc."""
    if pd.isna(cell_value):
        return None
    
    cell_str = str(cell_value).strip()
    if not cell_str or cell_str.upper() in ["D", "NA", "NAN", "NONE", "", "DECLINE"]:
        return None
    
    matches = re.findall(r'(\d+\.?\d*)%?', cell_str)
    valid_nums = []
    for m in matches:
        try:
            num = float(m)
            if 0 < num < 1:
                num *= 100
            valid_nums.append(num)
        except:
            continue
    
    return min(valid_nums) if valid_nums else None

def safe_float(value) -> Optional[float]:
    """Safely convert value to float"""
    if pd.isna(value):
        return None
    val_str = str(value).strip().upper()
    if val_str in ["D", "NA", "", "NAN", "NONE", "DECLINE"]:
        return None
    try:
        num = float(val_str.replace('%', '').strip())
        if 0 < num < 1:
            num = num * 100
        return num
    except:
        return None

def parse_age_based_values(cell_str: str) -> Dict[str, float]:
    """
    Parse age-based values like:
    'Age 0: 27.5%\nAge 1+: 26%'
    'Age 0: 27.5% Age 1+: 26%'
    'Age 0: 28%/27.5%   Age 1+: 26%'
    Returns: {'Age 0': 27.5, 'Age 1+': 26.0}
    """
    if not cell_str:
        return {}
    
    cell_str = str(cell_str)
    age_values = {}
    
    # Pattern for Age 0
    age0_match = re.search(r'Age\s*0\s*[:\-]?\s*([0-9\.\%/]+)', cell_str, re.IGNORECASE)
    if age0_match:
        val = extract_lowest_payin(age0_match.group(1))
        if val is not None:
            age_values['Age 0'] = val
    
    # Pattern for Age 1+
    age1_match = re.search(r'Age\s*1\s*\+\s*[:\-]?\s*([0-9\.\%/]+)', cell_str, re.IGNORECASE)
    if age1_match:
        val = extract_lowest_payin(age1_match.group(1))
        if val is not None:
            age_values['Age 1+'] = val
    
    return age_values

# ===============================================================================
# BASE PROCESSOR CLASS
# ===============================================================================

class CVProcessor(ABC):
    """Base class for all CV pattern processors"""
    
    @abstractmethod
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        """Detect if this processor can handle the given sheet"""
        pass
    
    @abstractmethod
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """Process the sheet and return records"""
        pass
    
    @abstractmethod
    def get_pattern_name(self) -> str:
        """Return the pattern name for display"""
        pass

# ===============================================================================
# PATTERN 1: PROBUS/BASE PATTERN (Works for Jan-May Standard)
# ===============================================================================

class ProbusPatternProcessor(CVProcessor):
    """
    Handles standard Probus pattern with region columns and CD2
    Works for: Probus_Jan25, Feb25, Mar25, May25, Digit CV 21-02-2025
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        if df.shape[0] < 3:
            return False
        
        # Check for region mapping in row 0
        top_row = df.iloc[0].astype(str).str.upper()
        has_regions = any(state in ' '.join(top_row.tolist()) for state in ["DELHI", "MUMBAI", "GOOD", "ROM"])
        
        # Check for CD2 in row 2
        if df.shape[0] > 2:
            row2 = df.iloc[2].astype(str).str.upper()
            has_cd2 = "CD2" in ' '.join(row2.tolist())
            return has_regions and has_cd2
        
        return False
    
    def get_pattern_name(self) -> str:
        return "Probus Base Pattern (Jan-May Standard)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        records = []
        
        try:
            print(f"\n{'='*80}")
            print(f"Processing with Probus Base Pattern")
            print(f"{'='*80}")
            
            # Map regions from row 0
            region_map = {}
            for col in range(df.shape[1]):
                val = str(df.iloc[0, col]).strip() if pd.notna(df.iloc[0, col]) else ""
                if val:
                    region_map[col] = val
            
            print(f"✓ Found {len(region_map)} region mappings")
            
            # Find CD2 columns from row 2
            cd2_columns = []
            for col in range(df.shape[1]):
                if pd.notna(df.iloc[2, col]) and "CD2" in str(df.iloc[2, col]).upper():
                    cd2_columns.append(col)
            
            print(f"✓ Found {len(cd2_columns)} CD2 columns")
            
            # Process data from row 3 onwards
            for idx in range(3, len(df)):
                row = df.iloc[idx]
                cluster = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if not cluster:
                    continue
                
                segment = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
                age_info = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "All"
                make = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "All"
                
                for cd2_col in cd2_columns:
                    region = region_map.get(cd2_col, "UNKNOWN")
                    state = STATE_MAPPING.get(region, region.upper())
                    
                    cell_value = row.iloc[cd2_col]
                    if pd.isna(cell_value):
                        continue
                    
                    # Determine policy type from row 1
                    policy_header = str(df.iloc[1, cd2_col]).upper() if len(df) > 1 else ""
                    policy_type = "Comp" if "COMP" in policy_header else "TP"
                    
                    # Extract payin - handle multiple formats
                    cell_str = str(cell_value).strip()
                    
                    # Case 1: Age conditions
                    age_matches = re.finditer(r'Age\s*\d*\-?\d*\s*[:\-]\s*(\d+\.?\d*)%?|Age\s*\d*\+?\s*[:\-]\s*(\d+\.?\d*)%', cell_str, re.IGNORECASE)
                    payin_list = []
                    for match in age_matches:
                        for g in match.groups():
                            if g:
                                val = safe_float(g)
                                if val is not None:
                                    payin_list.append(val)
                    
                    # Case 2: X%/Y% format - take smaller
                    if not payin_list:
                        slash_matches = re.findall(r'(\d+\.?\d*)\s*%?\s*/\s*(\d+\.?\d*)%?', cell_str)
                        for a, b in slash_matches:
                            v1 = safe_float(a)
                            v2 = safe_float(b)
                            if v1 is not None and v2 is not None:
                                payin_list.append(min(v1, v2))
                    
                    # Case 3: Single percentage
                    if not payin_list:
                        single = safe_float(cell_value)
                        if single is not None:
                            payin_list.append(single)
                    
                    # Process each payin value
                    for payin in payin_list:
                        payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, payin)
                        
                        records.append({
                            "State": state,
                            "Location/Cluster": region,
                            "Original Segment": cluster,
                            "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                            "Segment": segment,
                            "Make": make,
                            "Age": age_info,
                            "LOB": "CV",
                            "Policy Type": policy_type,
                            "Payin (CD2)": f"{payin:.2f}%",
                            "Payin Category": get_payin_category(payin),
                            "Calculated Payout": f"{payout:.2f}%",
                            "Formula Used": formula,
                            "Rule Explanation": rule_exp,
                            "Pattern": "Probus Base"
                        })
            
            print(f"✓ Extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"Error in Probus pattern: {e}")
            traceback.print_exc()
            return []

# ===============================================================================
# PATTERN 2: APRIL PATTERN (With Addon / Without Addon)
# ===============================================================================

class AprilPatternProcessor(CVProcessor):
    """
    April 2025 pattern with 'With Addon' / 'Without Addon' format
    Works for: Digit 13-04-2025
    
    Structure:
    - Row 1: Optional title
    - Row 2: Headers (RTO Cluster | Segment | Make | CD1 | CD2)
    - Row 3+: Data with addon/without addon values
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        # Look for "RTO" and "CLUSTER" together in headers
        for i in range(min(10, len(df))):
            row_str = ' '.join(df.iloc[i].astype(str).str.upper())
            if "RTO" in row_str and "CLUSTER" in row_str and "CD" in row_str:
                # Check for addon pattern in data
                sample = df.iloc[i+1:i+5].to_string().upper()
                if "ADDON" in sample:
                    return True
        return False
    
    def get_pattern_name(self) -> str:
        return "April 2025 Pattern (With/Without Addon)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        records = []
        
        try:
            print(f"\n{'='*80}")
            print(f"Processing with April Pattern (Addon Format)")
            print(f"{'='*80}")
            
            # Find header row
            header_row = None
            rto_col = segment_col = make_col = cd1_col = cd2_col = None
            
            for i in range(min(10, len(df))):
                for j in range(df.shape[1]):
                    cell_val = str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else ""
                    
                    if "RTO" in cell_val and "CLUSTER" in cell_val:
                        header_row = i
                        rto_col = j
                    elif "SEGMENT" == cell_val and header_row == i:
                        segment_col = j
                    elif "MAKE" == cell_val and header_row == i:
                        make_col = j
                    elif "CD1" == cell_val and header_row == i:
                        cd1_col = j
                    elif "CD2" == cell_val and header_row == i:
                        cd2_col = j
            
            if header_row is None:
                print("Could not find header row")
                return []
            
            print(f"✓ Header at row {header_row + 1}")
            print(f"✓ Columns: RTO={rto_col}, Segment={segment_col}, Make={make_col}, CD1={cd1_col}, CD2={cd2_col}")
            
            # Determine policy type from title row
            policy_type = "Comp"
            if header_row > 0 and cd1_col is not None:
                title = str(df.iloc[header_row - 1, cd1_col]).upper() if pd.notna(df.iloc[header_row - 1, cd1_col]) else ""
                if "TP" in title:
                    policy_type = "TP"
            
            # Process data rows
            for idx in range(header_row + 1, len(df)):
                row = df.iloc[idx]
                
                rto_cluster = str(row.iloc[rto_col]).strip() if rto_col is not None and pd.notna(row.iloc[rto_col]) else ""
                if not rto_cluster:
                    continue
                
                segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
                make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
                state = STATE_MAPPING.get(rto_cluster, rto_cluster.upper())
                
                # Process CD1
                if cd1_col is not None and cd1_col < len(row):
                    cd1_cell = str(row.iloc[cd1_col]).strip() if pd.notna(row.iloc[cd1_col]) else ""
                    
                    if cd1_cell:
                        # Check for With/Without Addon
                        with_match = re.search(r'with\s+addon[:\s]*(\d+\.?\d*)%?', cd1_cell, re.IGNORECASE)
                        without_match = re.search(r'without\s+addon[:\s]*(\d+\.?\d*)%?', cd1_cell, re.IGNORECASE)
                        
                        for match, addon_type in [(with_match, "CD1 With Addon"), (without_match, "CD1 Without Addon")]:
                            if match:
                                val = safe_float(match.group(1))
                                if val is not None:
                                    payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, val)
                                    
                                    records.append({
                                        "State": state,
                                        "Location/Cluster": rto_cluster,
                                        "Original Segment": segment,
                                        "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                                        "Make": make,
                                        "LOB": "CV",
                                        "Policy Type": policy_type,
                                        "Payin (CD2)": f"{val:.2f}%",
                                        "Payin Category": get_payin_category(val),
                                        "Calculated Payout": f"{payout:.2f}%",
                                        "Formula Used": formula,
                                        "Rule Explanation": rule_exp,
                                        "Addon Type": addon_type,
                                        "Pattern": "April Addon"
                                    })
                        
                        # Plain value if no addon pattern
                        if not with_match and not without_match:
                            val = safe_float(cd1_cell)
                            if val is not None:
                                payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, val)
                                
                                records.append({
                                    "State": state,
                                    "Location/Cluster": rto_cluster,
                                    "Original Segment": segment,
                                    "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                                    "Make": make,
                                    "LOB": "CV",
                                    "Policy Type": policy_type,
                                    "Payin (CD2)": f"{val:.2f}%",
                                    "Payin Category": get_payin_category(val),
                                    "Calculated Payout": f"{payout:.2f}%",
                                    "Formula Used": formula,
                                    "Rule Explanation": rule_exp,
                                    "Addon Type": "CD1 Plain",
                                    "Pattern": "April Addon"
                                })
                
                # Process CD2 (same logic)
                if cd2_col is not None and cd2_col < len(row):
                    cd2_cell = str(row.iloc[cd2_col]).strip() if pd.notna(row.iloc[cd2_col]) else ""
                    
                    if cd2_cell:
                        with_match = re.search(r'with\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
                        without_match = re.search(r'without\s+addon[:\s]*(\d+\.?\d*)%?', cd2_cell, re.IGNORECASE)
                        
                        for match, addon_type in [(with_match, "CD2 With Addon"), (without_match, "CD2 Without Addon")]:
                            if match:
                                val = safe_float(match.group(1))
                                if val is not None:
                                    payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, val)
                                    
                                    records.append({
                                        "State": state,
                                        "Location/Cluster": rto_cluster,
                                        "Original Segment": segment,
                                        "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                                        "Make": make,
                                        "LOB": "CV",
                                        "Policy Type": policy_type,
                                        "Payin (CD2)": f"{val:.2f}%",
                                        "Payin Category": get_payin_category(val),
                                        "Calculated Payout": f"{payout:.2f}%",
                                        "Formula Used": formula,
                                        "Rule Explanation": rule_exp,
                                        "Addon Type": addon_type,
                                        "Pattern": "April Addon"
                                    })
                        
                        if not with_match and not without_match:
                            val = safe_float(cd2_cell)
                            if val is not None:
                                payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, val)
                                
                                records.append({
                                    "State": state,
                                    "Location/Cluster": rto_cluster,
                                    "Original Segment": segment,
                                    "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                                    "Make": make,
                                    "LOB": "CV",
                                    "Policy Type": policy_type,
                                    "Payin (CD2)": f"{val:.2f}%",
                                    "Payin Category": get_payin_category(val),
                                    "Calculated Payout": f"{payout:.2f}%",
                                    "Formula Used": formula,
                                    "Rule Explanation": rule_exp,
                                    "Addon Type": "CD2 Plain",
                                    "Pattern": "April Addon"
                                })
            
            print(f"✓ Extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"Error in April pattern: {e}")
            traceback.print_exc()
            return []

# ===============================================================================
# PATTERN 3: MAY PATTERN 1 (Age-Based CV)
# ===============================================================================

class MayPattern1Processor(CVProcessor):
    """
    May Pattern 1 - Age-based values (Age 0, Age 1+)
    Works for: DIGIT 13-05-2025 (CV Worksheet)
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        # Check for age-based patterns
        sample_str = df.head(20).to_string().upper()
        return "AGE 0" in sample_str and "AGE 1+" in sample_str
    
    def get_pattern_name(self) -> str:
        return "May Pattern 1 (Age-Based CV)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        records = []
        
        try:
            print(f"\n{'='*80}")
            print(f"Processing with May Pattern 1 (Age-Based)")
            print(f"{'='*80}")
            
            # Find header row
            header_row = None
            cluster_col = None
            
            for i in range(min(10, len(df))):
                row_str = ' '.join(df.iloc[i].astype(str).str.upper())
                if "CLUSTER" in row_str and "CD2" in row_str:
                    header_row = i
                    # Find cluster column
                    for j in range(df.shape[1]):
                        if "CLUSTER" in str(df.iloc[i, j]).upper():
                            cluster_col = j
                            break
                    break
            
            if header_row is None:
                print("Could not find header")
                return []
            
            print(f"✓ Header at row {header_row + 1}")
            
            # Find CD2 columns
            cd2_cols = []
            for j in range(df.shape[1]):
                if pd.notna(df.iloc[header_row, j]) and "CD2" in str(df.iloc[header_row, j]).upper():
                    cd2_cols.append(j)
            
            print(f"✓ Found {len(cd2_cols)} CD2 columns")
            
            # Process data
            for idx in range(header_row + 1, len(df)):
                row = df.iloc[idx]
                cluster = str(row.iloc[cluster_col]).strip() if cluster_col is not None and pd.notna(row.iloc[cluster_col]) else ""
                if not cluster:
                    continue
                
                state = STATE_MAPPING.get(cluster, cluster.upper())
                
                for cd2_col in cd2_cols:
                    cell_value = row.iloc[cd2_col]
                    if pd.isna(cell_value):
                        continue
                    
                    cell_str = str(cell_value)
                    age_values = parse_age_based_values(cell_str)
                    
                    for age_label, payin in age_values.items():
                        payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", "Comp", payin)
                        
                        records.append({
                            "State": state,
                            "Location/Cluster": cluster,
                            "Original Segment": f"CV - {age_label}",
                            "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                            "Age": age_label,
                            "LOB": "CV",
                            "Policy Type": "Comp",
                            "Payin (CD2)": f"{payin:.2f}%",
                            "Payin Category": get_payin_category(payin),
                            "Calculated Payout": f"{payout:.2f}%",
                            "Formula Used": formula,
                            "Rule Explanation": rule_exp,
                            "Pattern": "May Age-Based"
                        })
            
            print(f"✓ Extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"Error in May Pattern 1: {e}")
            traceback.print_exc()
            return []


# ===============================================================================
# PATTERN 4: MAY PATTERN 2 (HCV Multi-Header with Segment Groups)
# ===============================================================================

class MayPattern2Processor(CVProcessor):
    """
    May Pattern 2 - HCV with multi-level headers and segment groups
    Works for: DIGIT 13-05-2025 (HCV), DIGIT 20-05-2025, DIGIT 24-05-2025
    
    Features:
    - Three-level headers (top=segment groups, mid=COMP/SATP, bottom=columns)
    - Merged cells for segment groups (Non-Dumper, Tipper, etc.)
    - With/Without Addon support
    - Referral note handling
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        if df.shape[0] < 5:
            return False
        
        # Check for multi-level headers
        top_str = ' '.join(df.iloc[:3].to_string().split()).upper()
        return ("NON-DUMPER" in top_str or "TIPPER" in top_str or 
                "DUMPER" in top_str or ("SEGMENT" in top_str and "SATP" in top_str))
    
    def get_pattern_name(self) -> str:
        return "May Pattern 2 (HCV Multi-Header)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        records = []
        
        try:
            print(f"\n{'='*80}")
            print(f"Processing with May Pattern 2 (HCV Multi-Header)")
            print(f"{'='*80}")
            
            # Find headers
            bottom_header_row = mid_header_row = top_header_row = None
            
            for i in range(min(20, len(df))):
                row_vals = [str(df.iloc[i, j]).upper().strip() if pd.notna(df.iloc[i, j]) else "" for j in range(df.shape[1])]
                
                if any("CLUSTER" in v for v in row_vals):
                    bottom_header_row = i
                    mid_header_row = i - 1 if i > 0 else None
                    top_header_row = i - 2 if i >= 2 else i - 3 if i >= 3 else None
                    break
            
            if bottom_header_row is None:
                print("Could not find headers")
                return []
            
            print(f"✓ Headers: Top={top_header_row}, Mid={mid_header_row}, Bottom={bottom_header_row}")
            
            # Build segment group mapping (forward fill for merged cells)
            segment_group_map = {}
            current_segment = None
            
            if top_header_row is not None:
                for j in range(df.shape[1]):
                    cell_val = str(df.iloc[top_header_row, j]).strip() if pd.notna(df.iloc[top_header_row, j]) else ""
                    if cell_val:
                        current_segment = cell_val
                    segment_group_map[j] = current_segment or "Unknown Segment"
            
            # Debug: Print segment groups
            unique_segs = set(segment_group_map.values())
            print(f"✓ Detected segment groups: {', '.join(unique_segs)}")
            
            # Find CD2 columns
            cd2_columns = []
            cluster_col = segment_col = make_col = age_from_col = age_to_col = None
            
            for j in range(df.shape[1]):
                bottom_val = str(df.iloc[bottom_header_row, j]).upper().strip() if pd.notna(df.iloc[bottom_header_row, j]) else ""
                
                if "CLUSTER" in bottom_val:
                    cluster_col = j
                elif bottom_val == "SEGMENT":
                    segment_col = j
                elif bottom_val == "MAKE":
                    make_col = j
                elif "AGE FROM" in bottom_val:
                    age_from_col = j
                elif "AGE TO" in bottom_val:
                    age_to_col = j
                elif bottom_val == "CD2":
                    # Determine policy type from mid row
                    policy_type = "Comp"
                    if mid_header_row is not None and pd.notna(df.iloc[mid_header_row, j]):
                        mid_val = str(df.iloc[mid_header_row, j]).upper().strip()
                        if "SATP" in mid_val:
                            policy_type = "SATP"
                        elif "TP" in mid_val:
                            policy_type = "TP"
                    
                    segment_group = segment_group_map.get(j, "Unknown Segment")
                    cd2_columns.append((j, policy_type, segment_group))
            
            print(f"✓ Found {len(cd2_columns)} CD2 columns")
            
            # Process data rows
            data_start_row = bottom_header_row + 1
            
            for idx in range(data_start_row, len(df)):
                row = df.iloc[idx]
                
                cluster = str(row.iloc[cluster_col]).strip() if cluster_col is not None and pd.notna(row.iloc[cluster_col]) else ""
                if not cluster or cluster.lower() in ["", "nan"]:
                    continue
                
                segment = str(row.iloc[segment_col]).strip() if segment_col is not None and pd.notna(row.iloc[segment_col]) else ""
                make = str(row.iloc[make_col]).strip() if make_col is not None and pd.notna(row.iloc[make_col]) else "All"
                age_from = row.iloc[age_from_col] if age_from_col is not None and pd.notna(row.iloc[age_from_col]) else ""
                age_to = row.iloc[age_to_col] if age_to_col is not None and pd.notna(row.iloc[age_to_col]) else ""
                
                state = STATE_MAPPING.get(cluster, cluster.upper())
                
                for cd2_col_idx, policy_type, segment_group in cd2_columns:
                    cell_value = row.iloc[cd2_col_idx]
                    if pd.isna(cell_value):
                        continue
                    
                    raw_str = str(cell_value).strip()
                    
                    # Check for referral notes
                    if re.search(r"for.*0.*1.*age|refer|grids?.*to.*be.*refer|above.*rates?", raw_str, re.IGNORECASE):
                        # Search previous rows for value
                        found = False
                        for prev_idx in range(idx - 1, data_start_row - 1, -1):
                            prev_cell = df.iloc[prev_idx, cd2_col_idx]
                            val = extract_lowest_payin(prev_cell)
                            if val is not None:
                                payin = val
                                addon_type = "Referred from above"
                                found = True
                                break
                        if not found:
                            continue
                    else:
                        # Extract with/without addon
                        with_match = re.search(r'with\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
                        without_match = re.search(r'without\s+addon[:\s]*([^\n]+)', raw_str, re.IGNORECASE)
                        
                        if with_match:
                            payin = extract_lowest_payin(with_match.group(1))
                            addon_type = "With Addon (Lowest)"
                        elif without_match:
                            payin = extract_lowest_payin(without_match.group(1))
                            addon_type = "Without Addon"
                        else:
                            payin = extract_lowest_payin(cell_value)
                            addon_type = "Plain/Multiple (Lowest)"
                        
                        if payin is None:
                            continue
                    
                    payout, formula, rule_exp = calculate_payout_with_formula("CV", "All GVW & PCV 3W, GCV 3W", policy_type, payin)
                    
                    records.append({
                        "State": state,
                        "Location/Cluster": cluster,
                        "Original Segment": segment,
                        "Mapped Segment": "All GVW & PCV 3W, GCV 3W",
                        "Segment Group": segment_group,
                        "Make": make,
                        "Age From": str(age_from),
                        "Age To": str(age_to),
                        "LOB": "CV",
                        "Policy Type": policy_type,
                        "Payin (CD2)": f"{payin:.2f}%",
                        "Payin Category": get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%",
                        "Formula Used": formula,
                        "Rule Explanation": rule_exp,
                        "Addon Type": addon_type,
                        "Pattern": "May HCV Multi-Header"
                    })
            
            print(f"✓ Extracted {len(records)} records")
            return records
            
        except Exception as e:
            print(f"Error in May Pattern 2: {e}")
            traceback.print_exc()
            return []

# ===============================================================================
# PATTERN 5: JUNE PATTERN 1 (Similar to May Pattern 2 but June specifics)
# ===============================================================================

class JunePattern1Processor(CVProcessor):
    """
    June Pattern 1 - Similar structure to May Pattern 2
    Works for: Digit 19-06-2025, 20-06-2025, 24-06-2025, JUNE 2025 (HCV)
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        # Similar to May Pattern 2 but check for June-specific markers
        if df.shape[0] < 5:
            return False
        
        sample = df.head(20).to_string().upper()
        # June patterns often have these characteristics
        has_structure = ("CLUSTER" in sample and "CD2" in sample and "SEGMENT" in sample)
        has_age_pattern = ("AGE" in sample and ("FROM" in sample or "0" in sample))
        
        return has_structure and has_age_pattern
    
    def get_pattern_name(self) -> str:
        return "June Pattern 1 (Age & Segment Groups)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        # Reuse May Pattern 2 logic as it's very similar
        processor = MayPattern2Processor()
        records = processor.process(df, sheet_name)
        
        # Update pattern name in records
        for record in records:
            record["Pattern"] = "June Pattern 1"
        
        return records

# ===============================================================================
# PATTERN 6: JUNE PATTERN 2 (Simplified CV Worksheet)
# ===============================================================================

class JunePattern2Processor(CVProcessor):
    """
    June Pattern 2 - Simpler CV worksheet format
    Works for: Digit JUNE 2025 (CV Worksheet - Pending)
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        # Simpler pattern - standard headers without complex multi-level
        if df.shape[0] < 3:
            return False
        
        sample = df.head(10).to_string().upper()
        return ("CLUSTER" in sample and "CD2" in sample and 
                "SEGMENT" not in sample)  # Distinguish from Pattern 1
    
    def get_pattern_name(self) -> str:
        return "June Pattern 2 (Simple CV)"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        # Use Probus pattern as base
        processor = ProbusPatternProcessor()
        records = processor.process(df, sheet_name)
        
        # Update pattern name
        for record in records:
            record["Pattern"] = "June Pattern 2"
        
        return records

# ===============================================================================
# PATTERN 7: JULY PATTERN 1
# ===============================================================================

class JulyPattern1Processor(CVProcessor):
    """
    July Pattern 1
    Works for: Digit 07-07-2025, DIGIT 18-07-2025
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        # Check for July-specific structure
        if df.shape[0] < 5:
            return False
        
        sample = df.head(15).to_string().upper()
        # July has specific header arrangement
        return ("CLUSTER" in sample and "CD2" in sample and 
                ("COMP" in sample or "SATP" in sample))
    
    def get_pattern_name(self) -> str:
        return "July Pattern 1"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        # Similar to May Pattern 2
        processor = MayPattern2Processor()
        records = processor.process(df, sheet_name)
        
        for record in records:
            record["Pattern"] = "July Pattern 1"
        
        return records

# ===============================================================================
# PATTERN 8: JULY PATTERN 2 / AUGUST PATTERN 1 (Combined)
# ===============================================================================

class JulyAugustPatternProcessor(CVProcessor):
    """
    July Pattern 2 / August Pattern 1 (same structure)
    Works for: PROBUS_Aug25 (HCV Worksheet)
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        if df.shape[0] < 5:
            return False
        
        sample = df.head(20).to_string().upper()
        # August/July pattern 2 has specific identifiers
        return ("PROBUS" in sample or "AUG" in sheet_name.upper() or
                "JULY" in sample)
    
    def get_pattern_name(self) -> str:
        return "July Pattern 2 / August Pattern 1"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        # Use May Pattern 2 as base
        processor = MayPattern2Processor()
        records = processor.process(df, sheet_name)
        
        for record in records:
            record["Pattern"] = "July/August"
        
        return records

# ===============================================================================
# PATTERN 9: SEPTEMBER PATTERN 1
# ===============================================================================

class SeptemberPattern1Processor(CVProcessor):
    """
    September Pattern 1
    Works for: September 2025 files
    """
    
    def detect(self, df: pd.DataFrame, sheet_name: str) -> bool:
        if df.shape[0] < 5:
            return False
        
        sample = df.head(15).to_string().upper()
        return "SEPT" in sheet_name.upper() or "SEP" in sheet_name.upper()
    
    def get_pattern_name(self) -> str:
        return "September Pattern 1"
    
    def process(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        # Use May Pattern 2 as base
        processor = MayPattern2Processor()
        records = processor.process(df, sheet_name)
        
        for record in records:
            record["Pattern"] = "September Pattern 1"
        
        return records

# ===============================================================================
# PATTERN DETECTOR AND DISPATCHER
# ===============================================================================

class CVPatternDetector:
    """Detects which CV pattern to use and routes to appropriate processor"""
    
    # All available processors in priority order (most specific first)
    PROCESSORS = [
        MayPattern2Processor(),           # Most complex - check first
        AprilPatternProcessor(),          # Addon-specific
        MayPattern1Processor(),           # Age-based
        JunePattern1Processor(),          # June with segments
        JunePattern2Processor(),          # June simple
        JulyPattern1Processor(),          # July specific
        JulyAugustPatternProcessor(),     # July/August combined
        SeptemberPattern1Processor(),     # September
        ProbusPatternProcessor(),         # Base/fallback
    ]
    
    @staticmethod
    def detect_pattern(df: pd.DataFrame, sheet_name: str) -> CVProcessor:
        """Detect and return the appropriate processor"""
        
        print(f"\n{'='*80}")
        print(f"🔍 Pattern Detection Process")
        print(f"{'='*80}")
        print(f"Sheet: {sheet_name}")
        print(f"Dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
        print(f"{'='*80}\n")
        
        for i, processor in enumerate(CVPatternDetector.PROCESSORS, 1):
            print(f"Testing {i}/{len(CVPatternDetector.PROCESSORS)}: {processor.get_pattern_name()}...", end=" ")
            
            try:
                if processor.detect(df, sheet_name):
                    print("✅ MATCH!")
                    print(f"\n{'='*80}")
                    print(f"✓ Selected Pattern: {processor.get_pattern_name()}")
                    print(f"{'='*80}\n")
                    return processor
                else:
                    print("❌ No match")
            except Exception as e:
                print(f"⚠️ Error: {e}")
                continue
        
        # Default to Probus pattern
        print(f"\n⚠️  No specific pattern detected")
        print(f"✓ Using default: Probus Base Pattern")
        print(f"{'='*80}\n")
        return ProbusPatternProcessor()
    
    @staticmethod
    def process_sheet(df: pd.DataFrame, sheet_name: str) -> Tuple[List[Dict], str]:
        """Detect pattern and process sheet"""
        processor = CVPatternDetector.detect_pattern(df, sheet_name)
        records = processor.process(df, sheet_name)
        return records, processor.get_pattern_name()

# ===============================================================================
# API ENDPOINTS
# ===============================================================================

@app.get("/")
async def root():
    return {
        "message": "DIGIT CV Unified Processor API - Complete Edition",
        "version": "2.0",
        "patterns_supported": [
            "Probus Base (Jan-May)",
            "April Addon Pattern",
            "May Pattern 1 (Age-Based)",
            "May Pattern 2 (HCV Multi-Header)",
            "June Pattern 1 & 2",
            "July Pattern 1 & 2",
            "August Pattern 1",
            "September Pattern 1"
        ]
    }

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel file and return available worksheets"""
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
        content = await file.read()
        xls = pd.ExcelFile(io.BytesIO(content))
        sheets = xls.sheet_names
        
        file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        uploaded_files[file_id] = {
            "content": content,
            "filename": file.filename,
            "sheets": sheets
        }
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "sheets": sheets,
            "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

@app.post("/process")
async def process_sheet(
    file_id: str,
    sheet_name: str
):
    """Process a specific worksheet with automatic pattern detection"""
    try:
        if file_id not in uploaded_files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
        file_data = uploaded_files[file_id]
        content = file_data["content"]
        
        if sheet_name not in file_data["sheets"]:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
        # Load sheet without header
        df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        print(f"\n{'#'*80}")
        print(f"# Processing Request")
        print(f"# File: {file_data['filename']}")
        print(f"# Sheet: {sheet_name}")
        print(f"{'#'*80}\n")
        
        # Detect pattern and process
        records, pattern_name = CVPatternDetector.process_sheet(df, sheet_name)
        
        if not records:
            return {
                "success": False,
                "message": "No records extracted. Please check the sheet structure.",
                "records": [],
                "count": 0,
                "pattern": pattern_name
            }
        
        # Calculate statistics
        states = {}
        policies = {}
        patterns = {}
        payins = []
        payouts = []
        
        for record in records:
            state = record.get("State", "Unknown")
            states[state] = states.get(state, 0) + 1
            
            policy = record.get("Policy Type", "Unknown")
            policies[policy] = policies.get(policy, 0) + 1
            
            pattern = record.get("Pattern", "Unknown")
            patterns[pattern] = patterns.get(pattern, 0) + 1
            
            try:
                payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
                payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
                if payin > 0:
                    payins.append(payin)
                    payouts.append(payout)
            except:
                pass
        
        avg_payin = sum(payins) / len(payins) if payins else 0
        avg_payout = sum(payouts) / len(payouts) if payouts else 0
        
        summary = {
            "total_records": len(records),
            "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
            "policies": policies,
            "patterns_used": patterns,
            "average_payin": round(avg_payin, 2),
            "average_payout": round(avg_payout, 2),
            "pattern": pattern_name
        }
        
        print(f"\n{'='*80}")
        print(f"✅ Processing Complete!")
        print(f"{'='*80}")
        print(f"Pattern Used: {pattern_name}")
        print(f"Total Records: {len(records)}")
        print(f"Average Payin: {avg_payin:.2f}%")
        print(f"Average Payout: {avg_payout:.2f}%")
        print(f"{'='*80}\n")
        
        return {
            "success": True,
            "message": f"Successfully processed {len(records)} records using {pattern_name}",
            "records": records,
            "count": len(records),
            "summary": summary
        }
        
    except Exception as e:
        print(f"\n{'!'*80}")
        print(f"! ERROR OCCURRED")
        print(f"{'!'*80}")
        traceback.print_exc()
        print(f"{'!'*80}\n")
        raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")

@app.post("/export")
async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
    """Export processed records to Excel file"""
    try:
        if not records:
            raise HTTPException(status_code=400, detail="No records to export")
        
        df = pd.DataFrame(records)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"CV_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, filename)
        
        df.to_excel(output_path, index=False, sheet_name='Processed')
        
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*80)
    print(" "*20 + "DIGIT CV UNIFIED PROCESSOR")
    print("="*80)
    print("\nSupported Patterns:")
    print("  ✓ Probus Base (Jan-May Standard)")
    print("  ✓ April Pattern (With/Without Addon)")
    print("  ✓ May Pattern 1 (Age-Based CV)")
    print("  ✓ May Pattern 2 (HCV Multi-Header)")
    print("  ✓ June Pattern 1 & 2")
    print("  ✓ July Pattern 1 & 2")
    print("  ✓ August Pattern 1")
    print("  ✓ September Pattern 1")
    print("="*80)
    print("\nStarting server on http://0.0.0.0:8000")
    print("="*80 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)
