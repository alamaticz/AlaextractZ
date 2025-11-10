import os
import re
import fitz  
import pandas as pd

# Helper: Clean text lines

def clean_lines(text):
    lines = text.split("\n")
    return [ln.strip() for ln in lines if ln.strip()]

# Extract common fields from PDF
def extract_common_fields(lines, text):
    common_data = {
        "SB Date": "",
        "SB NO": "",
        "Consignee Name": "",
        "Inv No": "",
        "Currency": "",
        "Exchange Rate": ""
    }

    # SB Date 
    m = re.search(r"(\d{1,2}-[A-Z]{3}-\d{2,4})", text)
    if m:
        common_data["SB Date"] = m.group(1)

    # SB Number (7-8 digit number)
    m = re.search(r"\b(\d{7,8})\b", text)
    if m:
        common_data["SB NO"] = m.group(1)

    # Invoice Number (e.g., JT-086/24-25)
    m = re.search(r"(JT[-/A-Z0-9]+)", text)
    if m:
        common_data["Inv No"] = m.group(1)

    # Consignee Name (appears below "Consignee")
    for i, ln in enumerate(lines):
        if "CONSIGNEE" in ln.upper():
            if i+1 < len(lines):
                common_data["Consignee Name"] = lines[i+1]
            break

    # Currency (appears after "4.CURRENC" label or in invoice section)
    # Look for currency code (SGD, USD, EUR, AED, INR, MYR, GBP)
    for i, ln in enumerate(lines):
        if "CURRENC" in ln.upper() or "INVOICE" in ln.upper():
            # Look in next 10 lines for a currency code
            for j in range(i+1, min(i+10, len(lines))):
                line = lines[j].strip()
                if re.match(r'^(SGD|USD|EUR|GBP|AED|MYR|INR|JPY|CNY|AUD|CAD|CHF)$', line):
                    common_data["Currency"] = line
                    break
            if common_data["Currency"]:
                break
    
    # Fallback: Extract from exchange rate line
    if not common_data["Currency"]:
        m = re.search(r"1\s+(SGD|USD|EUR|GBP|AED|MYR|INR|JPY|CNY|AUD|CAD|CHF)\s+INR", text)
        if m:
            common_data["Currency"] = m.group(1)

    # Exchange Rate (e.g., 1 SGD INR 60.8 or 1 USD INR 86.2)
    # Look for pattern with "EXCHANGE RATE" label nearby for better accuracy
    m = re.search(r"EXCHANGE\s+RATE.*?1\s+(?:SGD|USD|EUR|GBP|AED|MYR)\s+(?:INR)?\s*([\d\.]+)", text, re.DOTALL)
    if m:
        common_data["Exchange Rate"] = m.group(1)
    else:
        # Fallback: Look for the pattern anywhere (less accurate)
        m = re.search(r"1\s+(?:SGD|USD|EUR|GBP|AED|MYR)\s+INR\s+([\d\.]+)", text)
        if m:
            common_data["Exchange Rate"] = m.group(1)
    
    return common_data

# --------------------------
# Extract all items from PDF
# --------------------------
def extract_all_items(lines, text):
    items = []
    seen_items = set()  # To track duplicates

    # Pattern 1: Find all items in horizontal format (same line)
    # "06039000 FRESH MIXED FLOWERS & GARLANDS 632 KGS 3 1896"
    # Search in INVOICE DETAILS section (Part II) which has complete descriptions
    invoice_section_start = -1
    item_section_start = -1
    
    for i, ln in enumerate(lines):
        if "PART - II - INVOICE DETAILS" in ln or ("INVOICE" in ln and "DETAILS" in ln):
            invoice_section_start = i
        if "PART - III - ITEM DETAILS" in ln or (item_section_start < 0 and "ITEM DETAILS" in ln):
            item_section_start = i
            break
    
    # Search in INVOICE section if found, otherwise use ITEM DETAILS, otherwise full text
    search_text = text
    if invoice_section_start > 0 and item_section_start > invoice_section_start:
        # Search only in INVOICE section (between Part II and Part III)
        search_lines = lines[invoice_section_start:item_section_start]
        search_text = " ".join(search_lines)
    elif item_section_start > 0:
        # Fallback to ITEM DETAILS section
        search_lines = lines[item_section_start:min(item_section_start + 100, len(lines))]
        search_text = " ".join(search_lines)
    
    item_pattern = r"\b(\d{8})\s+([A-Z][A-Z\s&/,.\-]+?)\s+(\d+)\s+(KGS|NOS|PCS|MTR|LTR|UNT|BOX)\s+(\d+)\s+(\d+)"
    matches = re.finditer(item_pattern, search_text)
    
    for m in matches:
        # Clean up the description - remove extra spaces and trailing symbols
        desc = re.sub(r'\s+', ' ', m.group(2).strip())
        desc = re.sub(r'[&\s]+$', '', desc).strip()
        
        # Skip if description is too short or incomplete (likely a fragment)
        if len(desc) < 10 or desc.count(' ') < 1:
            continue
        
        # Create unique key to avoid duplicates
        item_key = (m.group(1), m.group(3), m.group(5), m.group(6))  # HS code, Qty, Rate, Total
        
        if item_key not in seen_items:
            seen_items.add(item_key)
            item = {
                "Description": desc,
                "Qty": m.group(3),
                "Rate": m.group(5),
                "Total": m.group(6)
            }
            items.append(item)
    
    # Pattern 2: Multi-line format (vertical columns) - ONLY if no items found yet
    if not items:
        # Search for descriptions in INVOICE section first, then ITEM DETAILS if needed
        search_range_start = 0
        search_range_end = len(lines)
        
        if invoice_section_start > 0 and item_section_start > invoice_section_start:
            # Expand range to before section header (descriptions often appear before the header)
            search_range_start = max(0, invoice_section_start - 50)
            search_range_end = item_section_start
        elif item_section_start > 0:
            # Fallback to ITEM DETAILS section
            search_range_start = max(0, item_section_start - 50)
            search_range_end = min(item_section_start + 100, len(lines))
        
        # Find all descriptions in multi-line format
        descriptions = []
        desc_indices = []
        hs_code_indices = set()  # Track which HS codes we've processed
        
        # First, find all HS codes and their descriptions in the search range
        for i in range(search_range_start, search_range_end):
            ln = lines[i]
            if re.match(r'^\d{8}$', ln.strip()) and i not in hs_code_indices:
                hs_code_indices.add(i)
                # Look for description in next few lines
                for j in range(i+1, min(i+10, search_range_end)):
                    line_text = lines[j].strip()
                    # Check if this looks like a description
                    if (line_text and 
                        re.match(r'^[A-Z]', line_text) and
                        not re.match(r'^\d+$', line_text) and
                        not re.match(r'^\d+\s+(KGS|NOS|PCS)', line_text) and
                        len(line_text) > 5 and
                        line_text not in descriptions):  # Avoid duplicate descriptions
                        descriptions.append(line_text)
                        desc_indices.append(j)
                        break
        
        # For multi-line format, extract Qty, Rate, Total for all items
        if descriptions:
            # Find the start position to search for quantities
            search_start = desc_indices[0] if desc_indices else 0
            
            # Collect all quantities, rates, and totals in vertical format
            quantities = []
            rates = []
            totals = []
            
            for i in range(search_start, min(search_start + 30, len(lines))):
                line = lines[i].strip()
                # Collect quantities (2-4 digits)
                if re.match(r'^\d{2,4}$', line) and len(quantities) < len(descriptions):
                    # Check if next few lines have KGS or similar unit
                    for j in range(i+1, min(i+5, len(lines))):
                        if re.match(r'^(KGS|NOS|PCS|MTR|LTR|UNT|BOX)$', lines[j].strip()):
                            quantities.append(line)
                            break
                
                # Collect rates (1-2 digits, after we have some quantities)
                if re.match(r'^\d{1,2}$', line) and len(quantities) > len(rates) and len(rates) < len(descriptions):
                    rates.append(line)
                
                # Collect totals (3-5 digits, after we have rates)
                if re.match(r'^\d{3,5}$', line) and len(rates) > len(totals) and len(totals) < len(descriptions):
                    totals.append(line)
            
            # Create items from collected data
            for idx, desc in enumerate(descriptions):
                item = {
                    "Description": desc,
                    "Qty": quantities[idx] if idx < len(quantities) else "",
                    "Rate": rates[idx] if idx < len(rates) else "",
                    "Total": totals[idx] if idx < len(totals) else ""
                }
                items.append(item)
    
    # If still no items found, try fallback methods
    if not items:
        item = {
            "Description": "",
            "Qty": "",
            "Rate": "",
            "Total": ""
        }
        
        # Try extracting Description using fallback method
        m_desc = re.search(r"\b(\d{8})\s+([A-Z][A-Z\s&/,.\-]+?)(?=\s+\d+\s+(?:KGS|NOS|PCS|MTR|LTR))", text)
        if m_desc:
            desc = re.sub(r'\s+', ' ', m_desc.group(2).strip())
            desc = re.sub(r'[&\s]+$', '', desc).strip()
            item["Description"] = desc
        
        # Qty, Rate, Total - look for first occurrence
        m_qty = re.search(r"(\d+)\s+(KGS|NOS|PCS|MTR|LTR|UNT|BOX)\s+(\d+)\s+(\d+)", text)
        if m_qty:
            item["Qty"] = m_qty.group(1)
            item["Rate"] = m_qty.group(3)
            item["Total"] = m_qty.group(4)
        
        items.append(item)
    
    return items

# --------------------------
# Process all PDFs in folder
# --------------------------
def process_shipping_bills(folder):
    rows = []
    for file in os.listdir(folder):
        if file.lower().endswith(".pdf"):
            path = os.path.join(folder, file)
            print(f"Processing → {file}")

            doc = fitz.open(path)
            text = ""
            for page in doc:
                text += page.get_text()
            doc.close()

            lines = clean_lines(text)
            
            # Extract common fields (same for all items in this PDF)
            common_data = extract_common_fields(lines, text)
            
            # Extract all items from this PDF
            items = extract_all_items(lines, text)
            
            # Create a row for each item
            for item in items:
                # Combine currency with rate (e.g., "SGD 3" or "USD 3")
                rate_with_currency = f"{common_data['Currency']} {item['Rate']}" if common_data['Currency'] and item['Rate'] else item['Rate']
                
                row = {
                    "SB Date": common_data["SB Date"],
                    "SB NO": common_data["SB NO"],
                    "Consignee Name": common_data["Consignee Name"],
                    "Inv No": common_data["Inv No"],
                    "Description": item["Description"],
                    "Qty": item["Qty"],
                    "Rate": rate_with_currency,
                    "Total": item["Total"],
                    "Exchange Rate": common_data["Exchange Rate"]
                }
                rows.append(row)

    return pd.DataFrame(rows)


# --------------------------
# MAIN EXECUTION
# --------------------------
if __name__ == "__main__":
    input_folder = "pdfs"  # <-- Folder inside project
    output_csv = "shipping_bill_output.csv"
    output_excel = "shipping_bill_output.xlsx"

    df = process_shipping_bills(input_folder)
    df.to_csv(output_csv, index=False)
    df.to_excel(output_excel, index=False)

    print("\n✅ Extraction Complete!")
    print(f"➡ CSV Saved:   {output_csv}")
    print(f"➡ Excel Saved: {output_excel}")
