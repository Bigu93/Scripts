import os
import argparse
from PIL import Image


def change_quality(source_folder, output_folder, quality=70):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    files = os.listdir(source_folder)

    image_files = [
        file
        for file in files
        if file.lower().endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp"))
    ]

    total_images = len(image_files)
    processed_images = 0

    for filename in image_files:
        try:
            image_path = os.path.join(source_folder, filename)
            img = Image.open(image_path)

            output_path = os.path.join(output_folder, filename)
            img.save(output_path, quality=quality)

            processed_images += 1
            progress_percentage = (processed_images / total_images) * 100
            print(f"\rProcessing... {int(progress_percentage)}%", end="", flush=True)

        except Exception as e:
            print(f"\nError processing '{filename}': {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Change quality of images in a folder")
    parser.add_argument(
        "source_folder", type=str, help="Input folder containing images"
    )
    parser.add_argument(
        "quality", type=int, default=70, nargs="?", help="Image quality (0-100)"
    )
    args = parser.parse_args()

    output_folder = os.path.join(
        args.source_folder, args.source_folder.split(os.path.sep)[-1] + "_READY"
    )

    change_quality(args.source_folder, output_folder, args.quality)

    print("\nEnd.")
