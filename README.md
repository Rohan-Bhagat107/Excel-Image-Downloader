# Excel-Image-Downloader
A Python script to download images from URLs listed in up to 2 specified or auto-detected columns of Excel files in a folder. It downloads images concurrently, handles hidden columns, retries failed downloads, and organizes images into separate folders named after each Excel file.

Excel Image Downloader is a Python script designed to automate the process of downloading images from URLs specified in Excel files. It supports reading multiple Excel files from a folder, detecting or using user-specified columns containing image URLs, and downloading images concurrently into organized folders.

Key Features

Input: Takes a directory containing Excel (.xlsx) files.

URL Extraction: Automatically detects up to 2 columns containing image URLs, or uses user-specified column names.

Image Downloading: Downloads images concurrently for faster performance with retries and integrity checks.

Output: Saves downloaded images into individual folders named after each Excel file, with counts of images downloaded.

Supports various image formats such as JPG, PNG, GIF, BMP, WEBP, and more.

Handles hidden Excel columns and only processes visible data.

Resilient downloads with retry mechanism and partial download cleanup.

Usage

Provide the folder path containing your Excel files.

Optionally specify the Excel column names that contain the image URLs (comma-separated).

Specify the output folder where images will be saved.

The script processes Excel files with names containing "te-IN.xlsx" and downloads images accordingly.
