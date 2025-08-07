import re
import PyPDF2
import fitz  # PyMuPDF
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import List, Optional
from docx import Document
import os

@dataclass
class ChainEntry:
    date: datetime
    date_string: str
    grantor: str
    grantee: str
    instrument: str
    book_page: str
    remark: str = ""
    is_vesting: bool = False
    line: str = ""

def parse_chain_from_pdf_tables(pdf_path: str) -> List[ChainEntry]:
    """Extracts chain of title entries from tables within a PDF."""
    entries = []
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Failed to open PDF {pdf_path}: {e}")
        return entries

    for page in doc:
        tables = page.find_tables()
        for table in tables:
            # Extract data as a list of lists
            table_data = table.extract()

            if not table_data:
                continue

            # Find header row and map columns
            header = [str(h).upper().replace('\n', ' ') for h in table_data[0]]
            try:
                # Find column indices, being flexible with naming
                date_idx = -1
                for d_col in ['DATED', 'DATE', 'FILED']:
                    if d_col in header:
                        date_idx = header.index(d_col)
                        break

                grantor_idx = header.index('GRANTOR')
                grantee_idx = header.index('GRANTEE')
                instrument_idx = header.index('INSTRUMENT')

                recording_idx = -1
                # More robust check for recording column header
                for r_col in ['RECORDING', 'BOOK-PAGE', 'BOOK_PAGE', 'BOOK/PAGE', 'BK/PG']:
                     if r_col in header:
                        recording_idx = header.index(r_col)
                        break

                if date_idx == -1 or recording_idx == -1 or grantor_idx == -1 or grantee_idx == -1 or instrument_idx == -1:
                    # If we can't find essential columns, skip this table
                    continue

            except ValueError:
                # If essential headers are missing, skip this table
                continue

            # Process data rows
            for row_data in table_data[1:]:
                # Check for malformed rows
                if len(row_data) != len(header):
                    continue

                date_str = (row_data[date_idx] or "").replace('\n', ' ').strip()
                grantor = (row_data[grantor_idx] or "").replace('\n', ' ').strip()
                grantee = (row_data[grantee_idx] or "").replace('\n', ' ').strip()
                instrument = (row_data[instrument_idx] or "").replace('\n', ' ').strip()
                book_page = (row_data[recording_idx] or "").replace('\n', ' ').strip()

                # Basic validation
                if not date_str or not book_page or (not grantor and not grantee):
                    continue

                parsed_date = parse_date(date_str)
                if parsed_date:
                    entry = ChainEntry(
                        date=parsed_date,
                        date_string=date_str,
                        grantor=grantor,
                        grantee=grantee,
                        instrument=instrument,
                        book_page=book_page,
                        is_vesting=is_vesting_deed(instrument),
                        line=' | '.join(map(str, row_data))
                    )
                    entries.append(entry)

    doc.close()
    return entries

def parse_date(date_str: str) -> Optional[datetime]:
    """Parse MM/DD/YYYY format date string."""
    try:
        parts = date_str.split('/')
        if len(parts) == 3:
            month = int(parts[0])
            day = int(parts[1])
            year = int(parts[2])
            return datetime(year, month, day)
    except (ValueError, IndexError):
        pass
    return None

def is_vesting_deed(instrument: str) -> bool:
    """Determine if an instrument is a vesting deed."""
    vesting_deed_types = [
        'WARRANTY DEED',
        'SPECIAL WARRANTY DEED',
        'QUITCLAIM DEED',
        'DEED',
        'GRANT DEED',
        'BARGAIN AND SALE DEED',
        'ASSUMPTION WARRANTY DEED',
        'CORRECTION DEED',
        'EXECUTOR\'S DEED',
        'ADMINISTRATOR\'S DEED',
        'TRUSTEE\'S DEED',
        'SHERIFF\'S DEED',
        'TAX DEED',
        'COMMISSIONER\'S DEED',
        'SPECIAL DEED',
        'BENEFICIARY DEED'
    ]

    non_vesting_types = [
        'DEED OF TRUST',
        'MORTGAGE',
        'ASSIGNMENT OF LEASES',
        'ASSIGNMENT OF RENTS',
        'ASSIGNMENT OF INCOME',
        'ASSIGNMENT OF LEASES, RENTS AND INCOME',
        'ASSIGNMENT OF LEASES, REENTS AND INCOME',  # Handle typos
        'UCC FINANCING STATEMENT',
        'SATISFACTION',
        'RELEASE',
        'SUBORDINATION',
        'MODIFICATION',
        'EXTENSION',
        'LIS PENDENS',
        'NOTICE OF DEFAULT',
        'AFFIDAVIT',
        'EASEMENT',
        'RIGHT OF WAY'
    ]

    upper_instrument = instrument.upper().strip()

    # First check if it's explicitly a non-vesting type
    for non_vesting in non_vesting_types:
        if non_vesting in upper_instrument:
            return False

    # Then check if it's a vesting type
    for vesting in vesting_deed_types:
        if vesting in upper_instrument:
            return True

    return False

def get_24_month_chain(entries: List[ChainEntry], processing_date: datetime = None) -> List[ChainEntry]:
    """
    Get chain of title covering AT LEAST the last 24 months.

    Logic:
    1. Find all vesting deeds
    2. Start from most recent and work backwards
    3. Include deeds until we cover at least 24 months from processing date
    4. If no deeds within 24 months, include the most recent vesting deed
    """
    if processing_date is None:
        processing_date = datetime.now()

    cutoff_date = processing_date - timedelta(days=730)  # 24 months ≈ 730 days

    # Get all vesting deeds, sorted by date (newest first)
    vesting_deeds = [entry for entry in entries if entry.is_vesting]
    vesting_deeds.sort(key=lambda x: x.date, reverse=True)

    if not vesting_deeds:
        return []

    result = []

    # If the most recent deed is within 24 months, include it and work backwards
    if vesting_deeds[0].date >= cutoff_date:
        # Start with the most recent deed
        result.append(vesting_deeds[0])

        # Work backwards to ensure we cover at least 24 months
        for deed in vesting_deeds[1:]:
            result.append(deed)

            # Check if we now cover at least 24 months
            earliest_date = min(deed.date for deed in result)
            if earliest_date <= cutoff_date:
                break
    else:
        # No deeds within 24 months, so include the most recent vesting deed
        result.append(vesting_deeds[0])

    # Sort result by date (newest first for display)
    result.sort(key=lambda x: x.date, reverse=True)
    return result

def create_title_document(chain_deeds: List[ChainEntry], output_path: str, template_path: str = None) -> bool:
    """Create title document with 24-month chain."""
    try:
        if template_path and os.path.exists(template_path):
            # Use provided template - this preserves the entire document
            doc = Document(template_path)
        else:
            # Create basic document
            doc = Document()
            doc.add_heading('TWENTY-FOUR MONTH CHAIN OF TITLE', level=1)

        # Find the chain table - look for any table that has chain-related headers
        chain_table = None

        # Search all tables for one with chain-related headers
        for table in doc.tables:
            if len(table.rows) > 0:
                header_row = table.rows[0]
                header_text = ' '.join([cell.text.upper() for cell in header_row.cells])

                # Look for typical chain table headers
                if any(keyword in header_text for keyword in ['GRANTOR', 'GRANTEE', 'INSTRUMENT', 'DATED', 'RECORDING']):
                    chain_table = table
                    break

        # If no suitable table found, look for tables near "TWENTY-FOUR MONTH CHAIN" text
        if not chain_table:
            for paragraph in doc.paragraphs:
                if "TWENTY-FOUR MONTH CHAIN" in paragraph.text.upper() or "24 MONTH CHAIN" in paragraph.text.upper():
                    # Look for tables after this paragraph
                    para_element = paragraph._element
                    parent = para_element.getparent()

                    # Find this paragraph's position and look for subsequent tables
                    found_para = False
                    for element in parent:
                        if found_para and element.tag.endswith('tbl'):
                            # Found a table after the chain paragraph
                            for table in doc.tables:
                                if table._element == element:
                                    chain_table = table
                                    break
                            break
                        if element == para_element:
                            found_para = True

                    if chain_table:
                        break

        # Create table if still not found (fallback)
        if not chain_table:
            chain_table = doc.add_table(rows=1, cols=5)
            chain_table.style = 'Table Grid'

            # Set headers
            header_cells = chain_table.rows[0].cells
            header_cells[0].text = 'GRANTOR:'
            header_cells[1].text = 'GRANTEE:'
            header_cells[2].text = 'INSTRUMENT'
            header_cells[3].text = 'DATED:'
            header_cells[4].text = 'RECORDING:'

        # Clear existing data rows (preserve header if it exists)
        rows_to_remove = []
        for i, row in enumerate(chain_table.rows):
            if i == 0:  # Keep first row as header
                continue
            rows_to_remove.append(row)

        # Remove data rows
        for row in rows_to_remove:
            chain_table._element.remove(row._element)

        # Add chain data
        if chain_deeds:
            for deed in chain_deeds:
                row = chain_table.add_row()
                cells = row.cells
                if len(cells) >= 5:
                    cells[0].text = deed.grantor.upper()
                    cells[1].text = deed.grantee.upper()
                    cells[2].text = deed.instrument.upper()
                    cells[3].text = deed.date_string
                    cells[4].text = deed.book_page
                else:
                    # Handle tables with different column counts
                    for i, text in enumerate([deed.grantor.upper(), deed.grantee.upper(), deed.instrument.upper(), deed.date_string, deed.book_page]):
                        if i < len(cells):
                            cells[i].text = text
        else:
            # Add empty row if no deeds found
            row = chain_table.add_row()
            if len(row.cells) > 0:
                row.cells[0].text = "No vesting deeds found"

        # Save the complete document (preserving all template content)
        doc.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating document: {e}")
        return False

def process_title_pdf(pdf_path: str, output_path: str, template_path: str = None) -> tuple[bool, str, List[ChainEntry]]:
    """
    Main function to process title PDF and create document.
    Returns: (success, message, chain_deeds)
    """
    try:
        # Parse entries from PDF tables
        entries = parse_chain_from_pdf_tables(pdf_path)
        if not entries:
            return False, "No chain of title entries found in PDF tables", []

        # Get 24-month chain
        chain_deeds = get_24_month_chain(entries)

        # Create document
        success = create_title_document(chain_deeds, output_path, template_path)

        if success:
            message = f"Document created successfully!\nFound {len(chain_deeds)} vesting deeds in 24-month chain."
            return True, message, chain_deeds
        else:
            return False, "Failed to create document", chain_deeds

    except Exception as e:
        return False, f"Error processing PDF: {str(e)}", []
