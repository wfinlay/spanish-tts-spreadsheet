# spanish-tts-spreadsheet
Automated Spanish text-to-speech generator that processes Excel/CSV files and adds audio 
file paths to adjacent columns. Supports multiple TTS engines including Google TTS, 
pyttsx3, macOS native voices and Azure Speech Services.   

**This script provides four different methods for generating Spanish audio files from your 
spreadsheet:**      

## Installation Requirements   

Be smart and use a [python virtual environment](https://realpython.com/python-virtual-environments-a-primer/).

	project_path="${HOME}/Documents/Source/Projects/spanish-tts-spreadsheet" && mkdir -p ${project_path} && cd ${project_path}
	python3 -m venv venv/
	source ./venv/bin/activate
	pip install -r requirements.txt 

## Repository Structure:

	spanish-tts-spreadsheet/
	├── spanish_tts_generator.py
	├── requirements.txt
	├── README.md
	├── LICENSE
	├── .gitignore
	├── examples/
	│   └── sample_spanish_words.xlsx
	└── docs/
	└── usage_examples.md

## Additional Virtual Environment Benefits:

1. **Reproducible environments** - others can recreate your exact setup
2. **Dependency isolation** - won't conflict with other Python projects
3. **Easy deployment** - clear dependency management



## Methods Available:   

gTTS (Google Text-to-Speech) - Recommended   

- High-quality Spanish voices   
- Requires internet connection   
- Free to use
- Supports multiple Spanish variants (Spain, Mexico, Argentina, etc.)


pyttsx3 - Offline option   

- Works without internet
- Limited Spanish voice quality (depends on your system)
- Generates WAV files

macOS Method Advantages:

- High Quality - Native macOS voices are very natural
- Offline - No internet required
- Fast - Native system integration
- Multiple Accents - Spain, Mexico, Argentina voices available
- Free - Built into macOS
- Adjustable Speed - Control speaking rate
- [See note below on audio format AIFF to MP3 conversion](#audio-format-notes)

Azure Speech Services - Premium option   

- Highest quality voices
- Requires paid API subscription
- Many Spanish accent options

## How to Use:  

See ```./docs/usage_examples.md``` for examples   

## Help

	python spanish_tts_generator.py --help

## Key Features:

**Automatic filename generation:**   
- Creates clean filenames based on the text content
- Cell-by-cell processing: Handles every non-empty cell in your spreadsheet
- Skip existing files: Won't regenerate files that already exist
- Error handling: Continues processing even if some cells fail
- Progress tracking: Shows which files are being generated

**File Naming:**   
Files are named as: ```row1_col1_hola_como_estas.mp3```   

## macOS Features:

**Available Spanish Voices on macOS:**

- Monica - Spanish (Spain) - Female voice
- Diego - Spanish (Argentina) - Male voice
- Jorge - Spanish (Spain) - Male voice
- Juan - Spanish (Mexico) - Male voice
- Paulina - Spanish (Mexico) - Female voice
- And more depending on your macOS version

Check Available Voices:

	python spanish_tts_generator.py --list-voices

This will show all Spanish voices available on your Mac.   

<a name="audio-format-notes"></a>
## Audio Format Notes

macOS say command outputs AIFF format by default   
The script will try to convert to MP3 if you have ffmpeg installed   
If no ffmpeg, it saves as AIFF (which works fine for most purposes)   

To Install ffmpeg (optional, for MP3 conversion):

	# Using Homebrew
	brew install ffmpeg

	# Or using MacPorts
	sudo port install ffmpeg


