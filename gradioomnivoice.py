import os
import shutil
import time
from gradio_client import Client

# Initialize the client with your updated active URL
client = Client("https://f3ebfce6867a004011.gradio.live/")

# 1. Define the exact text from your prompt
text_to_synthesize = (
    "Welcome to PNB Growth Plus Scheme for MSME Borrowers. PNB has revamped its Trade Growth Scheme into the PNB Growth Plus Scheme to boost credit access for MSMEs, especially micro enterprises, youth, and women entrepreneurs.The scheme now features simplified procedures and exemption from financial papers."
)

language = "English"
gender_choice = "Female / 女"
accent_choice = "Indian Accent / 印度口音"

# 2. Strictly verified Dropdown Arrays with perfect capitalization
age_variants = [
    "Auto", 
    "Child / 儿童",       
    "Teenager / 少年",
    "Young Adult / 青年",
    "Middle-aged / 中年",
    "Elderly / 老年"
]

pitch_variants = [
    "Auto", 
    "Very Low Pitch / 极低音调",
    "Low Pitch / 低音调",
    "Moderate Pitch / 中音调",
    "High Pitch / 高音调",
    "Very High Pitch / 极高音调"
]

# Create an output directory
output_dir = "generated_speech_variants"
os.makedirs(output_dir, exist_ok=True)

print("Starting audio generation loop with strict capitalization formatting...\n")
global_start_time = time.time()

# 3. Loop through all combinations of Age and Pitch
for age in age_variants:
    for pitch in pitch_variants:
        
        # Clean up names for clean, standard English filenames
        file_age = age.split(" / ")[0].lower().replace(" ", "_")
        file_pitch = pitch.split(" / ")[0].lower().replace(" ", "_")
        
        output_filename = f"english-pnb-female-{file_age}-{file_pitch}-1.wav"
        final_path = os.path.join(output_dir, output_filename)
        
        print(f"Generating: {output_filename}...")
        single_start_time = time.time()
        
        try:
            # Call the Voice Design endpoint
            result = client.predict(
                text=text_to_synthesize,
                lang=language,
                ns=32,            
                gs=2,
                dn=True,
                sp=0.9,
                du=0,             # Set to 0 matching your payload specifications
                pp=True,
                po=True,
                param_9=gender_choice,  
                param_10=age,           
                param_11=pitch,         
                param_12="Auto",        
                param_13=accent_choice, 
                param_14="Auto",        
                api_name="/_design_fn"
            )
            
            # Handle tuple/list responses safely
            if isinstance(result, (tuple, list)):
                temp_audio_path = result[0]
            else:
                temp_audio_path = result
            
            # Copy and save file if valid
            if temp_audio_path and isinstance(temp_audio_path, str) and os.path.exists(temp_audio_path):
                shutil.copy(temp_audio_path, final_path)
                generation_time = time.time() - single_start_time
                print(f"-> Successfully saved to: {final_path} (Took {generation_time:.2f} seconds)\n")
            else:
                print(f"-> Error: No valid file path returned for {output_filename}. Got: {result}\n")
                
        except Exception as e:
            print(f"-> Failed to generate {output_filename}. Error: {e}\n")

total_duration = time.time() - global_start_time
print("=" * 50)
print(f"All processing complete! Total time: {total_duration:.2f} seconds.")
print(f"Check your tracks inside the '{output_dir}' directory.")
print("=" * 50)