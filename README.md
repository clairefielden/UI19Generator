
<h1 align="center">
  <br>
  <img src="documents/LOGO.JPG" alt="UI-19 Generator" width="200"></a>
  <br>
  UI-19 Generator
  <br>
</h1>

<h4 align="center">A minimal UI Form Editor that generates UI-19s, Salary schedule and UI-2.8.</h4>

<p align="center">
  <a href="#key-features">Key Features</a> •
  <a href="#how-to-install">How To Install</a> •
  <a href="#how-to-run">How to Run</a> •
  <a href="#enhancement">Future Enhancement</a> •
  <a href="#credits">Credits</a> •
  <a href="#support">Support</a> •
  <a href="#license">License</a>
</p>

## Key Features

* Employee Information
  - All calculations are done manually
  - __First Name__: The first names(s) of the employee
  - __Surname__: The surname(s) of employee
  - __Clock Number__: The clock numer of the employee, as per the payroll system
  - __Remuneration__: The total salary (monthly remuneration)
  - __Hours worked__: The total number of hours worked in a month
* Termination Information
  - __Reason for termination__: Please do not make a spelling mistake, as your entry is used to determine the code for the new UI-19
  - __Commencement date__: The official start date of the employee, as per the payroll system
  - __Termination date__: The official date of termination of contract
* Company Information
  - __Pre-existing Database__: Every time you enter company information, it is saved in a database. Next time you need to generate a UI-19 for the same company, enter only the __UIF Number__, and all other information will automatically be loaded
  - __PAYE, UIF, CIPC Number__: The official numbers assigned to the company to which the employee belongs
  - __Company Name__: The full name, with Capital letters and lower case letters, no spelling mistakes, and "Pty/Ltd" where necessary
  - __Company Address__: The address of the company. If this is not filled in, the system will automatically set the address to be that of Prestige Payrolls 

## How To Install

To clone and run this application, you'll need [Git](https://git-scm.com) and [Powershell](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scripts?view=powershell-7.4) (which comes with [Windows](https://learn.microsoft.com/en-us/powershell/scripting/windows-powershell/starting-windows-powershell?view=powershell-7.4)) installed on your computer. From your command line:

```bash
# Clone this repository
$ git clone https://github.com/clairefielden/UI19Generator
```
Go to the folder where the directory is saved. \
Change the following filepaths in ```main.py``` to the **fully-qualified** paths of the directories where you want the UI-19s to be saved.
```python 
ORIGINAL_NEW_UI19
ORIGINAL_OLD_UI19
ORIGINAL_UI_28
ORIGINAL_SALARY_SCHED
NEW_PATH
```
Run the ```UI-19_Filler_Install.ps1``` script by right-clicking on it and selecting "Run with Powershell" \
If you get an error, install [Python](https://www.python.org/downloads/) 

Navigate to ```UI-19_Filler.ps```
Right-click on the script and selct "Open with notepad"
Change _directory of installation_ to the *fully-qualified* path of your UI-19 Filler installation

## How To Run

* Navigate to the folder ```UI19Filler``` on your desktop
* Right-click on the script ```UI-19_Filler.ps1```
* Select "Run with Powershell"

> **Note**
> If you're using **Powershell** for Windows, you can open the terminal as follows:
> Enter the following commands in the terminal:
```powershell
> cd cd c:\Users\FLDCLA001\Desktop\2024\UI19Filler
> dir
> Set-ExecutionPolicy RemoteSigned
> python main.py
```

## Future Enhancement

Import CSV of employee information and generate employee UI-19s without having to type by pressing the **Import** button.
> Please note this button is not yet impemented!

## Credits

This software uses the following open source packages:

- [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/)
- [Frontend](https://pypi.org/project/frontend/)
- [Fitz](https://github.com/chjj/marked)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)
- [Spire.Doc](https://pypi.org/project/Spire.Doc/)
- [Spire.PDF](https://pypi.org/project/Spire.Pdf/10.1.1/)

## Support

_accounts@prestigepayrolls.co.za_

## License

MIT

