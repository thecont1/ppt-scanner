# PowerPoint Slide Scanner

This friendly tool helps you quickly understand what's in your PowerPoint presentations by creating an Excel file with a neat overview of all your slides.

## What does it do?

- Scans through all PowerPoint files in a folder you choose
- Picks out the most important phrase from each slide
- Identifies slides that are just images
- Creates a nicely formatted Excel file with all this information

## Setup Instructions

1. Make sure you have Python installed on your computer
2. Create a file named `.env` in the same folder as the program and add your OpenAI API key like this:

   ```
   OPENAI_API_KEY=sk-Abc123xYz456MnP789QrStUvWxYz0123456789
   ```

3. Open Terminal (on Mac) or Command Prompt (on Windows)
4. Navigate to the program's folder
5. Install the required packages by typing:

   ```
   pip install -r requirements.txt
   ```

## How to Use

### The Easy Way (Drag and Drop)

1. Open Terminal
2. Type `python` (with a space after it)
3. Drag the `ppt-scanner.py` file into the Terminal window
4. Type another space
5. Drag the folder containing your PowerPoint files into the Terminal window
6. Press Enter

### Example

If your files are in a folder called "My Presentations", it might look like this:

```
python /Users/yourname/Desktop/ppt-scanner.py /Users/yourname/Documents/My\ Presentations
```

## What You Get

The program creates an Excel file named `ppt_slides.xlsx` in the same folder as your PowerPoint files. This Excel file will have:

- The name of each PowerPoint file
- The slide numbers
- A key phrase that tells you what's on each slide
- Clear marking of slides that are just images

## Tips

- The program works with both `.ppt` and `.pptx` files
- Image-only slides will be marked as "[Image Slide]"
- The Excel columns are sized for easy reading

## Need Help?

If you run into any issues or have questions, here are some common solutions:

1. Make sure your OpenAI API key is correctly set up in the `.env` file
2. Check that you have all the required packages installed
3. Make sure you're pointing to the correct folder containing your PowerPoint files

## Requirements

This program needs:

- Python 3.6 or newer
- The packages listed in `requirements.txt`
- An OpenAI API key
