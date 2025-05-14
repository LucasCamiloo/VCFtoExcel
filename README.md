# Contacts Extractor to Excel

This project extracts contacts from `.txt` files in vCard format, located in the `numeros` folder, and generates an Excel spreadsheet (`contatos_com_arquivo.xlsx`) with the following information:

- Contact name (decoding Quoted-printable names if necessary)
- Phone number (digits only)
- Source file name
- Indication if the number is repeated in other files

## How to use

1. **Prerequisites:**  
   Make sure you have [Node.js](https://nodejs.org/) installed.

2. **Install dependencies:**  
   In the terminal, run:
   ```
   npm install
   ```

3. **Add your contact files:**  
   Copy all `.txt` files containing contacts in vCard format to the `numeros` folder inside this directory. If your file is in `.vcf` format, you can simply rename the extension to `.txt`.

4. **Run the script:**  
   In the terminal, run:
   ```
   node mergeRepeatedNumbers.js
   ```

5. **Result:**  
   The file `contatos_com_arquivo.xlsx` will be generated in this folder, containing all extracted contacts and a column indicating if the number is repeated.

## Notes

- The script detects and decodes names encoded in Quoted-printable.
- Repeated phone numbers are marked in the "Repetido" column with "SIM".
- The source file name is included for easier identification.

---
Developed to help organize and analyze contacts extracted from backups or system exports.
