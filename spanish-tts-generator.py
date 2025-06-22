import pandas as pd
import os
from pathlib import Path
import re
import argparse
import sys
import subprocess
import platform
import pprint as pp
import re

# Method 1: Using gTTS (Google Text-to-Speech) - Free, requires internet
def generate_audio_gtts(text, filename, lang='es'):
    """Generate audio using Google Text-to-Speech"""
    try:
        from gtts import gTTS
        tts = gTTS(text=text, lang=lang, slow=False)
        tts.save(filename)
        return True
    except Exception as e:
        print(f"Error with gTTS for '{text}': {e}")
        return False

# Method 2: Using pyttsx3 - Offline, but may have limited Spanish voices
def generate_audio_pyttsx3(text, filename):
    """Generate audio using pyttsx3 (offline)"""
    try:
        import pyttsx3
        engine = pyttsx3.init()
        
        # Try to set Spanish voice (if available)
        voices = engine.getProperty('voices')
        for voice in voices:
            if 'spanish' in voice.name.lower() or 'es' in voice.id.lower():
                engine.setProperty('voice', voice.id)
                break
        
        # Set speech rate and volume
        engine.setProperty('rate', 150)  # Speed of speech
        engine.setProperty('volume', 0.9)  # Volume level (0.0 to 1.0)
        
        engine.save_to_file(text, filename)
        engine.runAndWait()
        return True
    except Exception as e:
        print(f"Error with pyttsx3 for '{text}': {e}")
        return False

# Method 3: Using Azure Cognitive Services (requires API key)
def generate_audio_azure(text, filename, subscription_key, region):
    """Generate audio using Azure Speech Services"""
    try:
        import azure.cognitiveservices.speech as speechsdk
        
        speech_config = speechsdk.SpeechConfig(subscription=subscription_key, region=region)
        speech_config.speech_synthesis_language = "es-ES"  # Spanish (Spain)
        # You can also use "es-MX" for Mexican Spanish, "es-AR" for Argentinian, etc.
        
        audio_config = speechsdk.audio.AudioOutputConfig(filename=filename)
        synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
        
        result = synthesizer.speak_text_async(text).get()
        return result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted
    except Exception as e:
        print(f"Error with Azure TTS for '{text}': {e}")
        return False

# Method 4: Using macOS built-in `say` command - High quality, offline
def generate_audio_macos_say(text, filename, voice='Monica', rate=200):
    """Generate audio using macOS built-in say command"""
    try:
        # Check if we're on macOS
        if platform.system() != 'Darwin':
            print("macOS say command is only available on macOS")
            return False
        
        # Convert mp3 to aiff (say command outputs AIFF)
        if filename.endswith('.mp3'):
            aiff_filename = filename.replace('.mp3', '.aiff')
        else:
            aiff_filename = filename
        
        # Use say command to generate audio
        cmd = [
            'say',
            '-v', voice,  # Voice name
            '-r', str(rate),  # Speaking rate (words per minute)
            '-o', aiff_filename,  # Output file
            text
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            # Convert AIFF to MP3 if needed (requires ffmpeg)
            if filename.endswith('.mp3') and aiff_filename != filename:
                try:
                    convert_cmd = ['ffmpeg', '-i', aiff_filename, '-y', filename]
                    convert_result = subprocess.run(convert_cmd, capture_output=True)
                    if convert_result.returncode == 0:
                        os.remove(aiff_filename)  # Remove temporary AIFF file
                    else:
                        # If ffmpeg fails, keep the AIFF file and update filename
                        print(f"Note: Saved as AIFF format (install ffmpeg for MP3): {aiff_filename}")
                        return True
                except FileNotFoundError:
                    print(f"Note: Saved as AIFF format (install ffmpeg for MP3): {aiff_filename}")
                    return True
            
            return True
        else:
            print(f"say command failed: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"Error with macOS say for '{text}': {e}")
        return False

def list_macos_spanish_voices():
    """List available Spanish voices on macOS"""
    try:
        if platform.system() != 'Darwin':
            return []
        
        result = subprocess.run(['say', '-v', '?'], capture_output=True, text=True)
        if result.returncode == 0:
            voices = []
            for line in result.stdout.split('\n'):
                if 'es_' in line or 'spanish' in line.lower():
                    voice_name = line.split()[0]
                    voices.append((voice_name, line.strip()))
            return voices
        return []
    except:
        return []

def generate_macos_say_commands(list_of_macos_voices):
    """ """
    say_string = "'¡Hola! Me llamo Reed.'"
    rate = 160
    regex = re.compile('(?P<voice>.*\)\))')
    for item in list_of_macos_voices:
        match = regex.match(item[1])
        if match:
            voice = match.groupdict()['voice']
            print(f"say -v '{voice}' -r {rate} '¡Hola! Me llamo {voice}.'")

def clean_filename(text, max_length=50):
    """Create a clean filename from text"""
    # Remove special characters and limit length
    clean_text = re.sub(r'[^\w\s-]', '', text).strip()
    clean_text = re.sub(r'[-\s]+', '_', clean_text)
    return clean_text[:max_length]

def process_spreadsheet_with_paths(file_path, spanish_column, output_dir='audio_files', method='gtts', save_updated_file=True):
    """
    Process spreadsheet and generate audio files, adding file paths to adjacent column
    
    Args:
        file_path: Path to Excel/CSV file
        spanish_column: Column name or index containing Spanish text (e.g., 'A', 0, 'Spanish_Words')
        output_dir: Directory to save audio files
        method: 'gtts', 'pyttsx3', or 'azure'
        save_updated_file: Whether to save the updated spreadsheet
    """
    
    # Create output directory
    Path(output_dir).mkdir(exist_ok=True)
    
    # Read spreadsheet
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:  # Excel file
        df = pd.read_excel(file_path)
    
    # Handle column specification (name or index)
    if isinstance(spanish_column, str):
        if spanish_column.isdigit():
            col_idx = int(spanish_column)
            spanish_col_name = df.columns[col_idx]
        else:
            spanish_col_name = spanish_column
            col_idx = df.columns.get_loc(spanish_column)
    else:  # numeric index
        col_idx = spanish_column
        spanish_col_name = df.columns[col_idx]
    
    # Create audio file path column name
    audio_col_name = f"{spanish_col_name}_Audio_Path"
    
    # Add the audio path column if it doesn't exist
    if audio_col_name not in df.columns:
        # Insert the new column right after the Spanish column
        df.insert(col_idx + 1, audio_col_name, '')
    
    successful = 0
    failed = 0
    
    # Process each row in the specified column
    for row_idx, spanish_text in enumerate(df[spanish_col_name]):
        if pd.notna(spanish_text) and str(spanish_text).strip():
            text = str(spanish_text).strip()
            
            # Create filename
            clean_text = clean_filename(text)
            filename = f"{output_dir}/row{row_idx+1}_{clean_text}.mp3"
            
            # Check if file already exists
            if os.path.exists(filename):
                print(f"Using existing file: {filename}")
                df.at[row_idx, audio_col_name] = filename
                successful += 1
                continue
            
            print(f"Generating audio for: '{text}' -> {filename}")
            
            # Generate audio based on selected method
            success = False
            
            if method == 'gtts':
                success = generate_audio_gtts(text, filename)
            elif method == 'pyttsx3':
                # pyttsx3 saves as .wav, so change extension
                filename = filename.replace('.mp3', '.wav')
                success = generate_audio_pyttsx3(text, filename)
            elif method == 'macos':
                success = generate_audio_macos_say(text, filename, MACOS_VOICE, MACOS_RATE)
            elif method == 'azure':
                # You'll need to provide your Azure credentials
                success = generate_audio_azure(text, filename, AZURE_SUBSCRIPTION_KEY, AZURE_REGION)
            
            if success:
                successful += 1
                df.at[row_idx, audio_col_name] = filename
                print(f"✓ Generated: {filename}")
            else:
                failed += 1
                df.at[row_idx, audio_col_name] = "FAILED"
                print(f"✗ Failed: {filename}")
    
    # Save the updated spreadsheet
    if save_updated_file:
        base_name = Path(file_path).stem
        extension = Path(file_path).suffix
        output_file = f"{base_name}_with_audio_paths{extension}"
        
        if extension.lower() == '.csv':
            df.to_csv(output_file, index=False)
        else:  # Excel
            df.to_excel(output_file, index=False)
        
        print(f"\nUpdated spreadsheet saved as: {output_file}")
    
    print(f"Completed! Successfully generated: {successful}, Failed: {failed}")
    return df

def process_spreadsheet(file_path, output_dir='audio_files', method='gtts'):
    """
    Legacy function - Process spreadsheet and generate audio files for all cells
    (kept for backward compatibility)
    """
    
    # Create output directory
    Path(output_dir).mkdir(exist_ok=True)
    
    # Read spreadsheet
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:  # Excel file
        df = pd.read_excel(file_path)
    
    successful = 0
    failed = 0
    
    # Process each row and column
    for row_idx, row in df.iterrows():
        for col_idx, cell_value in enumerate(row):
            if pd.notna(cell_value) and str(cell_value).strip():
                text = str(cell_value).strip()
                
                # Create filename
                clean_text = clean_filename(text)
                filename = f"{output_dir}/row{row_idx+1}_col{col_idx+1}_{clean_text}.mp3"
                
                # Skip if file already exists
                if os.path.exists(filename):
                    print(f"Skipping existing file: {filename}")
                    continue
                
                print(f"Generating audio for: '{text}' -> {filename}")
                
                # Generate audio based on selected method
                success = False
                
                if method == 'gtts':
                    success = generate_audio_gtts(text, filename)
                elif method == 'pyttsx3':
                    # pyttsx3 saves as .wav, so change extension
                    filename = filename.replace('.mp3', '.wav')
                    success = generate_audio_pyttsx3(text, filename)
                elif method == 'azure':
                    # You'll need to provide your Azure credentials
                    subscription_key = "YOUR_AZURE_SUBSCRIPTION_KEY"
                    region = "YOUR_AZURE_REGION"  # e.g., "eastus"
                    filename = filename.replace('.mp3', '.wav')
                    success = generate_audio_azure(text, filename, subscription_key, region)
                
                if success:
                    successful += 1
                    print(f"✓ Generated: {filename}")
                else:
                    failed += 1
                    print(f"✗ Failed: {filename}")
    
    print(f"\nCompleted! Successfully generated: {successful}, Failed: {failed}")



# Additional utility functions for advanced use cases
def batch_convert_specific_cells(file_path, cell_ranges, output_dir='selected_audio'):
    """Convert only specific cells/ranges (for advanced scripting)"""
    Path(output_dir).mkdir(exist_ok=True)
    
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    for cell_range in cell_ranges:
        # Example: cell_range could be "A1", "B2:B10", etc.
        # This is a simplified version - you'd need to parse Excel-style ranges
        pass

def convert_with_custom_voice_settings():
    """Example of using gTTS with different Spanish variants (for testing)"""
    text = "Hola, ¿cómo estás?"
    
    # Different Spanish variants
    variants = {
        'es': 'Spanish (Spain)',
        'es-mx': 'Spanish (Mexico)', 
        'es-ar': 'Spanish (Argentina)',
        'es-co': 'Spanish (Colombia)'
    }
    
    for lang_code, description in variants.items():
        try:
            from gtts import gTTS
            tts = gTTS(text=text, lang=lang_code, slow=False)
            tts.save(f"sample_{lang_code}.mp3")
            print(f"Generated sample for {description}")
        except:
            print(f"Language {lang_code} not supported")

def config_parser_parameters():
    parser = argparse.ArgumentParser(
        description='Generate Spanish audio files from spreadsheet and add file paths to adjacent column',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s my_spanish_words.xlsx
  %(prog)s data.csv --column "Spanish_Text" --method pyttsx3
  %(prog)s words.xlsx --column 0 --output-dir "./audio" --no-save
        """
    )
    
    # Required argument
    parser.add_argument(
        'file_path',
        help='Path to the Excel (.xlsx, .xls) or CSV file containing Spanish text'
    )
    
    # Optional arguments
    parser.add_argument(
        '-c', '--column',
        default=0,
        help='Spanish text column (name, index, or Excel letter). Default: 0 (first column)'
    )
    
    parser.add_argument(
        '-m', '--method',
        choices=['gtts', 'pyttsx3', 'azure', 'macos'],
        default='gtts',
        help='Text-to-speech method. Default: gtts'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        default='audio_files',
        help='Directory to save audio files. Default: audio_files'
    )
    
    parser.add_argument(
        '--no-save',
        action='store_true',
        help='Do not save updated spreadsheet with audio paths'
    )
    
    parser.add_argument(
        '--azure-key',
        help='Azure Speech Services subscription key (required for --method azure)'
    )
    
    parser.add_argument(
        '--macos-voice',
        default='Monica',
        help='macOS voice name for Spanish TTS (use --list-voices to see options). Default: Monica'
    )
    
    parser.add_argument(
        '--macos-rate',
        type=int,
        default=200,
        help='macOS speech rate in words per minute. Default: 200'
    )
    
    parser.add_argument(
        '--list-voices',
        action='store_true',
        help='List available Spanish voices on macOS and exit'
    )

    return parser.parse_args()



def main():
    """Main function to handle command line arguments and execute the script"""

    args = config_parser_parameters()
    # Validate file exists
    if not os.path.exists(args.file_path):
        print(f"Error: File '{args.file_path}' not found.")
        sys.exit(1)

    # Validate file extension
    valid_extensions = ['.xlsx', '.xls', '.csv']
    file_ext = Path(args.file_path).suffix.lower()
    if file_ext not in valid_extensions:
        print(f"Error: Unsupported file type '{file_ext}'. Supported: {', '.join(valid_extensions)}")
        sys.exit(1)

    # Handle column argument (convert string numbers to int)
    column = args.column
    if isinstance(column, str) and column.isdigit():
        column = int(column)
    
    # Validate macOS method
    if args.method == 'macos' and platform.system() != 'Darwin':
        print("Error: --method macos is only available on macOS")
        sys.exit(1)
    
    if args.method == 'list_voices':
        if platform.system() != 'Darwin':
            print("Error: --method macos is only available on macOS")
            sys.exit(1)
        else:
            list_of_macos_voices = list_macos_spanish_voices()
            generate_macos_say_commands(list_of_macos_voices)
            sys.exit(0)
    
    # Validate Azure credentials if needed
    if args.method == 'azure':
        if not args.azure_key or not args.azure_region:
            print("Error: --azure-key and --azure-region are required when using --method azure")
            sys.exit(1)
        # Update the global Azure credentials
        global AZURE_SUBSCRIPTION_KEY, AZURE_REGION, MACOS_VOICE, MACOS_RATE
        AZURE_SUBSCRIPTION_KEY = args.azure_key
        AZURE_REGION = args.azure_region
    
    # Set macOS voice settings
    MACOS_VOICE = args.macos_voice
    MACOS_RATE = args.macos_rate
    
    print(f"Processing file: {args.file_path}")
    print(f"Spanish column: {column}")
    print(f"TTS method: {args.method}")
    print(f"Output directory: {args.output_dir}")
    print("-" * 50)
    
    try:
        # Process the spreadsheet
        df = process_spreadsheet_with_paths(
            file_path=args.file_path,
            spanish_column=column,
            output_dir=args.output_dir,
            method=args.method,
            save_updated_file=not args.no_save
        )
        
        print("\n" + "="*50)
        print("SUCCESS: Processing completed!")
        
        if not args.no_save:
            base_name = Path(args.file_path).stem
            extension = Path(args.file_path).suffix
            output_file = f"{base_name}_with_audio_paths{extension}"
            print(f"Updated spreadsheet: {output_file}")
        
        print(f"Audio files saved in: {args.output_dir}/")
        
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        sys.exit(1)

# Global variables for Azure credentials and macOS settings (updated by command line args)
AZURE_SUBSCRIPTION_KEY = "YOUR_AZURE_SUBSCRIPTION_KEY"
AZURE_REGION = "YOUR_AZURE_REGION"
MACOS_VOICE = "Monica"
MACOS_RATE = 200


# Example usage
if __name__ == "__main__":
    main()