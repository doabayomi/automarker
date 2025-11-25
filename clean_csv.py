import csv
import tempfile
import os

def process_csv(csv_path):
    # Create a temporary file
    with tempfile.NamedTemporaryFile('w', newline='', encoding='utf-8', delete=False) as tmpfile:
        with open(csv_path, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.DictReader(infile)
            fieldnames = ['index'] + reader.fieldnames

            writer = csv.DictWriter(tmpfile, fieldnames=fieldnames)
            writer.writeheader()

            for i, row in enumerate(reader, start=1):
                processed_row = {'index': i}

                for key, value in row.items():
                    if isinstance(value, str):
                        # Apply namecase (capitalize first letter, rest lowercase)
                        value = value.strip().capitalize()

                    processed_row[key] = value

                writer.writerow(processed_row)

    # Replace the original file with the updated one
    os.replace(tmpfile.name, csv_path)
    print(f"File '{csv_path}' processed successfully.")

# Example usage:
# process_csv_in_place("data.csv")

if __name__ == "__main__":
    process_csv('submitters.csv')