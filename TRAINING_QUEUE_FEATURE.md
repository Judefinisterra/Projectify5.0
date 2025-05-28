# Training Data Queue Feature

## Overview
The "Add to Training Data Queue" feature allows developers to easily capture training data pairs consisting of user prompts and selected code snippets for later analysis and model training.

## How to Use

1. **Enter a prompt** in the message box (the main input field in developer mode)
2. **Select text** in the code editor (optional - you can highlight specific code you want to associate with the prompt)
3. **Click the "Add to Training Data Queue" button**

The system will:
- Capture the current prompt from the message box
- Capture any selected text from the code editor
- Create a simple entry with both pieces of data
- Save the entry to localStorage for persistence
- Automatically download a CSV file with all training data entries

## Data Structure

Each training data entry contains:
- `prompt`: The user's input prompt (empty string if not provided)
- `selectedCode`: The selected text from the code editor (empty string if none selected)

## CSV Format

The downloaded CSV file contains two columns (no header):
- Column 1: Prompt text
- Column 2: Selected code text

Example CSV content:
```
"Create a pizza sales model","<SPREAD-E; format=""Volume""; row1=""V1|Pizza Sales|||||100|200|300|"";>"
"Add revenue calculation","<MULT2-S; driver1=""V1""; driver2=""V2""; row1=""R1|Revenue|||||F|F|F|"";>"
```

## Storage

- Training data is stored in localStorage under the key `trainingDataQueue`
- Each time you click the button, a new CSV file is downloaded with ALL entries (not just the new one)
- The CSV filename includes the current date: `training_data_queue_YYYY-MM-DD.csv`

## Use Cases

1. **Prompt Engineering**: Capture effective prompts along with their expected code outputs
2. **Model Training**: Build datasets for fine-tuning AI models
3. **Documentation**: Keep track of common user requests and their solutions
4. **Quality Assurance**: Review prompt-code pairs for consistency and accuracy

## Technical Details

- The feature uses browser localStorage for persistence
- CSV generation handles proper escaping of quotes and newlines
- File download uses the browser's built-in download mechanism
- No server-side storage is required
- Simple two-column format: prompt, code

## Error Handling

- If neither prompt nor selected code is provided, an error message is shown
- If CSV download fails, the data is still saved to localStorage and logged to console
- Malformed localStorage data is handled gracefully with fallback to empty array 