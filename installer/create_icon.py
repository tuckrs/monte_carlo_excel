"""
Convert SVG icon to ICO format.
"""
import subprocess
import sys

def create_icon():
    """Convert SVG to ICO using ImageMagick."""
    try:
        # Install cairosvg for SVG handling
        subprocess.check_call([
            sys.executable,
            "-m",
            "pip",
            "install",
            "cairosvg",
            "pillow"
        ])
        
        from cairosvg import svg2png
        from PIL import Image
        import io
        
        # Convert SVG to PNG
        png_data = svg2png(url="installer_icon.svg")
        
        # Create ICO file
        img = Image.open(io.BytesIO(png_data))
        img.save("installer_icon.ico", format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        
        print("Icon created successfully!")
        
    except Exception as e:
        print(f"Error creating icon: {e}")
        sys.exit(1)

if __name__ == "__main__":
    create_icon()
