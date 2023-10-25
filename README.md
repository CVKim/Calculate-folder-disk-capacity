# Calculate-folder-disk-capacity
------------------------------
## Folder & Image Sizes Extractor

This script extracts sizes and allocated sizes of image files in a chosen directory and its subdirectories. Results can be saved in either `txt` or `excel` format.

## Table of Contents

- [Features](#features)
- [Usage](#usage)
- [Code Breakdown](#code-breakdown)
- [Note](#note)

## Features

* When a directory is chosen, it scans through all its subfolders.
* Calculates size and allocated size for `.jpg` and `.bmp` image files in each folder.
* Saves the results in either `txt` or `excel` format.

## Usage

1. Run the script.
2. Use the GUI to choose a directory.
3. Choose the format to save the results (`txt` or `excel`).

## Code Breakdown

* **`get_folder_and_image_sizes(path)`**: Extracts image files' sizes and allocated sizes for the given path.
* **`save_as_txt(folder_path, results)`**: Saves the results in a txt file format.
* **`save_as_excel(folder_path, results)`**: Saves the results in an excel file format.
* **`main()`**: The main execution function. It prompts the user to choose a directory and file format.

## Note

After selecting a directory and choosing `txt` as the format, for instance, the results will be saved in a `.txt` file inside the selected directory.
