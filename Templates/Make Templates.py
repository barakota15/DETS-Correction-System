from PIL import Image
from imutils.perspective import four_point_transform
from imutils import contours
import imutils
import cv2
import numpy as np
import argparse
import random


def show_images(images, titles, kill_later=True):
    for index, image in enumerate(images):
        cv2.imshow(titles[index], image)
    cv2.waitKey(0)
    if kill_later:
        cv2.destroyAllWindows()
        
        
def find_corners(image_path):
    img = Image.open(image_path)
    width, height = img.size

    left, top, right, bottom = width, height, 0, 0

    for x in range(width):
        for y in range(height):
            pixel = img.getpixel((x, y))
            if pixel == (0, 0, 0):  # Assuming the black squares have RGB values of (0, 0, 0)
                left = min(left, x)
                top = min(top, y)
                right = max(right, x)
                bottom = max(bottom, y)

    return left, top, right, bottom

def crop_image(image_path, output_path):
    left, top, right, bottom = find_corners(image_path)
    img = Image.open(image_path)
    img_cropped = img.crop((left, top, right, bottom))
    img_cropped.save(output_path)
    
    
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--image", default=output_path, help="path to the input image")
    args = vars(ap.parse_args())
    # Resize the image to a smaller size for processing
    ig = cv2.imread(args['image'])
    scale_percent = 50  # percent of original size
    width = int(ig.shape[1] * scale_percent / 100)
    height = int(ig.shape[0] * scale_percent / 100)
    dim = (width, height)
    image = cv2.resize(ig, dim, interpolation=cv2.INTER_AREA)
    cv2.imwrite("temp.png", image)
    show_images([image], ["image"])

    print(f"Image cropped and saved at {output_path}")

# Specify your input and output paths
input_image_path = "Template-60.png"
output_image_path = "output_cropped_image.png"

crop_image(input_image_path, output_image_path)
