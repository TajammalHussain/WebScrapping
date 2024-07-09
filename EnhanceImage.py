from PIL import Image, ImageEnhance

def enhance_image(input_path, output_path, scale_factor=2, enhancement_factor=1.5):
    # Open the original image
    image = Image.open(input_path)
    
    # Calculate new dimensions
    new_width = int(image.width * scale_factor)
    new_height = int(image.height * scale_factor)
    
    # Resize the image
    resized_image = image.resize((new_width, new_height), Image.LANCZOS)
    
    # Enhance the image (contrast, brightness, sharpness)
    enhancer = ImageEnhance.Contrast(resized_image)
    enhanced_image = enhancer.enhance(enhancement_factor)
    
    enhancer = ImageEnhance.Brightness(enhanced_image)
    enhanced_image = enhancer.enhance(enhancement_factor)
    
    enhancer = ImageEnhance.Sharpness(enhanced_image)
    enhanced_image = enhancer.enhance(enhancement_factor)
    
    # Save the enhanced image
    enhanced_image.save(output_path, format='PNG')
    print(f"Enhanced image saved to {output_path}")

# Example usage with the specified paths
input_image_path = r'C:\Users\Tajammal\Desktop\Mixed\logo-color-2.png'  # Ensure to include file extension if needed
output_image_path = r'C:\Users\Tajammal\Desktop\Mixed\EnhancedXpressLogo.png'  # Ensure to include file extension

enhance_image(input_image_path, output_image_path)
