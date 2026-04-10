from typing import Optional
import os
import tempfile
import io
import requests
import re
from PIL import Image

class ImageAgent:
    """Scrapes images from Unsplash based on keywords."""
    
    def __init__(self):
        self.output_dir = tempfile.mkdtemp(prefix="md2pptx_images_")
        
    def generate_image(self, prompt: str, index: int) -> Optional[str]:
        """Attempt to generate an image using nanao-bana-2, fallback to Unsplash."""
        if not prompt:
            return None
            
        print(f"    [ImageAgent] Attempting to generate image with gemini-3.1-flash-image-preview for: {prompt[:50]}...")
        try:
            from .client import get_client
            client = get_client()
            result = client.models.generate_images(
                model='gemini-3.1-flash-image-preview',
                prompt=prompt,
                config=dict(
                    number_of_images=1,
                    output_mime_type="image/jpeg",
                )
            )
            for generated_image in result.generated_images:
                image = Image.open(io.BytesIO(generated_image.image.image_bytes))
                filepath = os.path.join(self.output_dir, f"img_{index}.jpg")
                image.save(filepath)
                return filepath
        except Exception as e:
            print(f"    [ImageAgent] Warning: 'gemini-3.1-flash-image-preview' generation failed ({e}). Falling back to LoremFlickr...")
            
        print(f"    [ImageAgent] Sourcing fallback image for: {prompt[:50]}...")
        try:
            # We use loremflickr for a reliable semantic image fallback
            import urllib.parse
            # Use the first couple of nouns/adjectives from the prompt as keywords
            keywords = ",".join([w for w in prompt.split() if len(w) > 3][:3])
            safe_keywords = urllib.parse.quote(keywords)
            url = f"https://loremflickr.com/1024/768/{safe_keywords}"
            
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            }
            response = requests.get(url, headers=headers, timeout=15)
            
            if response.status_code == 200:
                image = Image.open(io.BytesIO(response.content))
                filepath = os.path.join(self.output_dir, f"img_{index}.jpg")
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                image.save(filepath)
                return filepath
            
            print(f"    [ImageAgent] Warning: LoremFlickr returned status {response.status_code}")
                
        except Exception as e:
            print(f"    [ImageAgent] Warning: Image sourcing failed. ({e})")
        return None
