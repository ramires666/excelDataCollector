import pandas as pd
import re
import chardet

# Define the regex pattern for license plates
# Matches 6-8 alphanumeric characters with at least one digit, no 4+ consecutive letters, excluding word-like patterns
pattern1 = r'[\w\d]{6,8}'
pattern2 = r'\b[a-zA-ZА-Яа-я]{2,11}\b'


def detect_encoding(file_path):
    """Detect the encoding of the file."""
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        return result['encoding']


def extract_plate(text):
    # Find all matches for the first pattern
    matches1 = re.findall(pattern1, text)

    if not matches1:
        return None

    # Process each match with the second pattern
    processed_matches = []
    for match in matches1:
        match2 = re.search(pattern2, match)
        if match2:
            # Remove the matched word-like portion
            processed = match.replace(match2.group(0), '')
            if processed:  # Only add if something remains after removal
                processed_matches.append(processed)
        else:
            processed_matches.append(match)

    # Return all processed matches joined by semicolon if there are multiple
    return '; '.join(processed_matches) if processed_matches else None


def parse_license_plates(input_file, output_file):
    try:
        # Detect the encoding of the input file
        encoding = detect_encoding(input_file)
        print(f"Detected encoding: {encoding}")

        # Read the text file into a list, stripping whitespace and ignoring empty lines
        with open(input_file, 'r', encoding=encoding) as file:
            lines = [line.strip() for line in file if line.strip()]

        # Create a DataFrame with two columns: 'original' and 'license_plate'
        df = pd.DataFrame({
            'original': lines,
            'license_plate': [None] * len(lines)  # Initialize with None
        })

        # Apply the parsing function to each row and update the 'license_plate' column
        df['license_plate'] = df['original'].apply(extract_plate)

        # Save the DataFrame to an Excel file
        df.to_excel(output_file, index=False)
        print(f"Successfully processed and saved to {output_file}")

    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Example usage
if __name__ == "__main__":
    input_file = r"C:\Users\delxps\PycharmProjects\excelCollector\mercenary_licence_plates2clean_wialon_270225.txt"
    output_file = "mercenaries_wialon_claened_license_plates.xlsx"

    parse_license_plates(input_file, output_file)