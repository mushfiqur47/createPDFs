# Create PDFs with Reportlab and Python!
Tired of creating a lot of pdfs with the same information only to change one parameter? 
With this file you can create n numbers of pdf certificates with a specific design using data from excel in seconds :)

This uses `reportlab` library for Python which creates a canvas to "draw" your document and creates a pdf file with that. 

## Ejecution
1. Install python
2. Install reportlab
  ```
  pip3 install pandas reportlab
  ```
4. Install openpyxl
  ```
  pip3 install openpyxl
  ```
4. Run the code
  ```
  python3 createPDF.py
  ```

It will create all the pdfs in the same folder

TODO: send each pdf to each email after creating them
