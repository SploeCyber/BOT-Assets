import os
from PIL import Image
import concurrent.futures
import time

def optimize_image(file_path):
    """
    Optimizes an image to reduce file size while maintaining high quality.
    Saves the optimized image with '_optimized.png' suffix.
    """
    try:
        # Check if the file is an image
        if not file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
            return

        # Check if it's already an optimized image to avoid re-optimizing
        if file_path.lower().endswith('_optimized.png'):
            return

        # Open the image
        with Image.open(file_path) as img:
            # Construct new filename
            base, ext = os.path.splitext(file_path)
            new_file_path = f"{base}_optimized.png"

            # Skip if optimized file already exists
            # if os.path.exists(new_file_path):
            #     print(f"Skipping (already optimized): {file_path}")
            #     return

            print(f"Optimizing: {file_path}")

            # Resize if necessary
            max_size = 1024
            if img.width > max_size or img.height > max_size:
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                print(f"Resized to: {img.size}")

            # Optimize and save
            img.save(new_file_path, 'PNG', optimize=True)
            
            print(f"Saved: {new_file_path}")

    except Exception as e:
        print(f"Error optimizing {file_path}: {e}")

# ... (optimize_image function remains the same) ...

def main():
    assets_dir = r'd:\GitHub\BOT-Assets\Assets'
    
    if not os.path.exists(assets_dir):
        print(f"Directory not found: {assets_dir}")
        return

    print(f"Scanning directory: {assets_dir}")
    
    files_to_process = []
    for root, dirs, files in os.walk(assets_dir):
        for file in files:
            file_path = os.path.join(root, file)
            files_to_process.append(file_path)

    print(f"Found {len(files_to_process)} files. Starting optimization with parallel processing...")
    
    start_time = time.time()
    
    # Use ProcessPoolExecutor for CPU-bound tasks like image processing
    # max_workers defaults to the number of processors on the machine
    with concurrent.futures.ProcessPoolExecutor() as executor:
        executor.map(optimize_image, files_to_process)

    end_time = time.time()
    print(f"Optimization completed in {end_time - start_time:.2f} seconds.")

if __name__ == "__main__":
    main()
