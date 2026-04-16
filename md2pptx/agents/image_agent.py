from typing import Optional
import os
from dotenv import load_dotenv
import tempfile
import io
import requests
import re
from PIL import Image

class ImageAgent:
    """Scrapes images from Unsplash based on keywords."""
    
    def __init__(self):
        dotenv_path = os.path.join(os.path.dirname(__file__), "..", ".env")
        load_dotenv(dotenv_path)
        self.output_dir = tempfile.mkdtemp(prefix="md2pptx_images_")
        self.unsplash_access_key = os.getenv("unsplash_access_key")
        
    def generate_image(self, prompt: str, index: int) -> Optional[str]:
        """Attempt to generate an image using nanao-bana-2, fallback to Unsplash."""
        if not prompt:
            return None
            
        # [ImageAgent] Commented out Gemini image generation as per request
        # print(f"    [ImageAgent] Attempting to generate image with imagen-3.0-generate-001 for: {prompt[:50]}...")
        # try:
        #     from .client import get_client
        #     client = get_client()
        #     result = client.models.generate_images(
        #         model='imagen-3.0-generate-001',
        #         prompt=prompt,
        #         config=dict(
        #             number_of_images=1,
        #             output_mime_type="image/jpeg",
        #         )
        #     )
        #     for generated_image in result.generated_images:
        #         image = Image.open(io.BytesIO(generated_image.image.image_bytes))
        #         filepath = os.path.join(self.output_dir, f"img_{index}.jpg")
        #         image.save(filepath)
        #         return filepath
        # except Exception as e:
        #     print(f"    [ImageAgent] Warning: 'imagen-3.0-generate-001' generation failed ({e}). Falling back to Unsplash/LoremFlickr...")

        # Attempt Unsplash API first
        if self.unsplash_access_key:
            print(f"    [ImageAgent] Sourcing image from Unsplash for: {prompt[:50]}...")
            try:
                import urllib.parse
                search_query = urllib.parse.quote(prompt[:100])
                url = f"https://api.unsplash.com/search/photos?query={search_query}&per_page=1&client_id={self.unsplash_access_key}"
                response = requests.get(url, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    if data.get('results'):
                        img_url = data['results'][0]['urls']['regular']
                        img_response = requests.get(img_url, timeout=15)
                        if img_response.status_code == 200:
                            image = Image.open(io.BytesIO(img_response.content))
                            filepath = os.path.join(self.output_dir, f"img_{index}.jpg")
                            if image.mode != 'RGB':
                                image = image.convert('RGB')
                            image.save(filepath)
                            return filepath
                print(f"    [ImageAgent] Unsplash API returned status {response.status_code} or no results.")
            except Exception as e:
                print(f"    [ImageAgent] Warning: Unsplash API sourcing failed ({e}).")
            
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

    def generate_multiple_images(self, prompts: list, base_index: int) -> list:
        """Generate images for multiple prompts, returning a list of valid paths."""
        results = []
        for offset, prompt in enumerate(prompts):
            path = self.generate_image(prompt, base_index * 10 + offset)
            if path:
                results.append(path)
        return results
