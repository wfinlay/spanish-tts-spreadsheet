# Usage Examples

**Note:** Make sure to activate your [python virtual environment](https://realpython.com/python-virtual-environments-a-primer/).

	project_path="${HOME}/Documents/Source/Projects/spanish-tts-spreadsheet" && cd ${project_path}
	source ./venv/bin/activate

## Basic Usage:

	# Simple usage - processes first column with default settings
	python spanish_tts_generator.py my_spanish_words.xlsx
	
	# Specify a different column by name
	python spanish_tts_generator.py data.csv --column "Spanish_Text"
	
	# Use a different TTS method
	python spanish_tts_generator.py words.xlsx --method pyttsx3
	
	# Custom output directory
	python spanish_tts_generator.py data.xlsx --output-dir "./my_audio_files"
	
	# Don't save updated spreadsheet (just generate audio)
	python spanish_tts_generator.py words.csv --no-save
	
## Advanced Usage:

	# Use Azure Speech Services
	python spanish_tts_generator.py data.xlsx --method azure --azure-key "your_key" --azure-region "eastus"
	
	# Specify column by index (0 = first column)
	python spanish_tts_generator.py words.xlsx --column 2
	
	# Full example with all options
	python spanish_tts_generator.py spanish_data.xlsx \
	    --column "Spanish_Words" \
	    --method gtts \
	    --output-dir "./audio_output" \
	    --no-save

	# Use macOS TTS with default Spanish voice (Monica)
	python spanish_tts_generator.py spanish_words.xlsx --method macos
	
	# List all available Spanish voices on your Mac
	python spanish_tts_generator.py --list-voices
	
	# Use a specific Spanish voice
	python spanish_tts_generator.py data.xlsx --method macos --macos-voice "Diego"
	
	# Adjust speaking rate (words per minute)
	python spanish_tts_generator.py words.csv --method macos --macos-voice "Paulina" --macos-rate 180
	
	# Full example with macOS options
	python spanish_tts_generator.py spanish_data.xlsx \
	    --method macos \
	    --macos-voice "Jorge" \
	    --macos-rate 220 \
	    --output-dir "./macos_audio"

## Command Line Options:

	file_path (required): Path to your Excel or CSV file
	--column / -c: Column containing Spanish text (name, index, or number)
	--method / -m: TTS method (gtts, pyttsx3, azure, macos)
	--output-dir / -o: Directory for audio files
	--no-save: Skip saving updated spreadsheet
	--azure-key: Azure subscription key (for Azure method)
	--azure-region: Azure region (for Azure method)

## Help:

	python spanish_tts_generator.py --help

This will show all available options and examples.

## Error Handling:

The script validates:

- File exists
- File format is supported (.xlsx, .xls, .csv)
- Azure credentials are provided when using Azure method
- Provides clear error messages and exits gracefully
