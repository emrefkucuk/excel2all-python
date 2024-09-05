# excel2all-python
# Turn Excel compatible files (i.e XLSX) into CSV, JSON, HTML, XML and PDF files

The code in this repository is licensed under the MIT License, which can be found in the [LICENSE](./LICENSE) file.
The included DejaVuSans.ttf font is licensed under the terms of the Bitstream Vera and Arev fonts, as detailed in the [FONT-LICENSE.md](./FONT-LICENSE.md) file.

I made this terminal based Python script as an assignment during my internship in [TÜRASAŞ](https://www.turasas.gov.tr/). The basic algorithm took roughly two days to complete but the conversions were a real hassle and took a week to perfect. I also wanted to try writing this according to the [PEP8](https://peps.python.org/pep-0008/) standards as best as I could, as a practice. Keep in mind that you can edit the ASCII banner in line 34 to write whatever you want according to your project.

## How it works
- After running the script, you will be greeted with a welcome screen: _Press `1` to choose an Excel file, press `2` to exit_
- Pressing `1` opens a dialog box prompting you to choose an Excel-compatible file (.xls, .xlsx, .xlsm, .xlsb).
- After choosing, you will be prompted to choose a directory where a new directory will be created with the same name as your Excel spreadsheet.
- After choosing, the program will run the necessary conversions. PDF conversion may take a while because it uses the `reportlab` library.
- After all the conversions are complete, you will get a message box informing you of the location of the new directory.
- Now, you will be back at the beginning. Procceed or exit as you would like.

## To-Do for me:
- Add English language support
- Allow more types of conversions (both from excel and from other formats)
- Add options for conversions (encoding, size, orientation, etc.)
- Fix text formatting and terminal colors
