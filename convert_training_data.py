import json
import os

def convert_to_jsonl():
    # File paths
    input_file = "../../../Dropbox/B - Freelance/C_Projectify/VanPC/Chain Reaction/Chain Reaction Model_6.4.25 v3.txt"
    output_file = "../../../Downloads/TrainingDataUTF8.jsonl"
    
    system_message = "You are a financial model planning assistant. You are tasked with programming the client's request with a custom financial model coding lanaguage you have been trained on"
    
    # Read the input file
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Skip the header line
    data_lines = lines[1:]
    
    # Read existing JSONL data
    existing_entries = []
    if os.path.exists(output_file):
        with open(output_file, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    existing_entries.append(json.loads(line.strip()))
    
    # Process each line and convert to JSONL format
    new_entries = []
    for line in data_lines:
        line = line.strip()
        if line and '@@@' in line:
            # Split on @@@ to get input and output
            parts = line.split('@@@', 1)
            if len(parts) == 2:
                user_input = parts[0].strip()
                assistant_output = parts[1].strip()
                
                # Remove quotes if they wrap the entire input
                if user_input.startswith('"') and user_input.endswith('"'):
                    user_input = user_input[1:-1]
                
                # Create the JSONL entry
                entry = {
                    "messages": [
                        {"role": "system", "content": system_message},
                        {"role": "user", "content": user_input},
                        {"role": "assistant", "content": assistant_output}
                    ]
                }
                new_entries.append(entry)
    
    # Combine existing and new entries
    all_entries = existing_entries + new_entries
    
    # Write to JSONL file
    with open(output_file, 'w', encoding='utf-8') as f:
        for entry in all_entries:
            f.write(json.dumps(entry) + '\n')
    
    print(f"Successfully converted {len(new_entries)} new training examples")
    print(f"Total entries in file: {len(all_entries)}")
    print(f"Output written to: {output_file}")

if __name__ == "__main__":
    convert_to_jsonl() 