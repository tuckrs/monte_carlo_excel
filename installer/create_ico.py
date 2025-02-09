from PIL import Image, ImageDraw

def create_icon():
    # Create images of different sizes
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    images = []
    
    for size in sizes:
        # Create a new image with a transparent background
        image = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # Calculate dimensions
        margin = size[0] // 8
        box_size = size[0] - (2 * margin)
        
        # Draw a green Excel-like square
        draw.rectangle(
            [margin, margin, margin + box_size, margin + box_size],
            fill='#107C41'  # Excel green
        )
        
        # Add to list
        images.append(image)
    
    # Save as ICO with multiple sizes
    images[0].save(
        'installer_icon.ico',
        format='ICO',
        append_images=images[1:],
        sizes=sizes
    )

if __name__ == "__main__":
    create_icon()
