
import argparse
import os
from PIL import Image

from apps.work.smr_ollama_extract import try_extract_smr_field_via_ollama

# Configuration for Ollama (adjust as needed)
OLLAMA_HTTP_BASE = "http://localhost:11434"  # Default Ollama API base URL
OLLAMA_MODEL = "llava"  # Default LLaVA model name

def main():
    parser = argparse.ArgumentParser(description="Process an image with a local AI model (Ollama/LLaVA).")
    parser.add_argument("image_path", type=str, help="Path to the input image file.")
    parser.add_argument("prompt", type=str, help="The prompt for the AI model.")
    args = parser.parse_args()

    if not os.path.exists(args.image_path):
        print(f"Error: Image file not found at {args.image_path}")
        return

    try:
        # Load the image using PIL
        img = Image.open(args.image_path).convert("RGB")
    except Exception as e:
        print(f"Error loading image: {e}")
        return

    print(f"Processing image '{args.image_path}' with prompt: '{args.prompt}'")

    try:
        # Call the Ollama extraction function
        result = try_extract_smr_field_via_ollama(
            img,
            args.prompt,  # The prompt is used as field_name here, which might need adjustment based on actual usage
            base_url=OLLAMA_HTTP_BASE,
            model=OLLAMA_MODEL,
        )

        if result.get("ok"):
            print("\nAI Response:")
            print(f"  Text: {result.get('text')}")
            print(f"  Confidence: {result.get('confidence')}")
            print(f"  Raw Model Output: {result.get('raw_model')}")
        else:
            print("\nAI processing failed:")
            print(f"  Reason: {result.get('reason')}")
            print(f"  Raw Model Output: {result.get('raw_model')}")

    except Exception as e:
        print(f"An error occurred during AI processing: {e}")

if __name__ == "__main__":
    main()
