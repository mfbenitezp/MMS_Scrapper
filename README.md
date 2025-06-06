# MMS WebScrapper

These are a set of Python scripts to scrape several charts and data from MMS for the exam board reports. This is part of the role of the Exam Officer.

## Requirements for the code to run:

1. You need the list of modules in a format like: "GG4258" no spaces
2. The links for the required charts (scatterplot with grade distribution and scatterplot with grade distribution from previous years) are https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG3214/Final+grade/GraphPage https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG3281/Final+grade/SubmitResults and the link for the grade table is https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG1002/Final+grade/ you will notice both links have the module code as a parameter
3. You can adapt the code to update the list of modules per semester and also the AY, as for now, the link is consistent across modules.
4. The main script is `ModuleGradesChartsExtractor.py`; the other scripts that describe parts of the process, but I kept them just for testing and adapting in the future.
5. Now the code in here just allows you to install the requirements in an independent Python environment. Once that is done, you can just open a terminal and run: python ModuleGradesChartsExtractor.py or python  module_charts_downloader.py

`# pip install requests beautifulsoup4 plotly kaleido selenium webdriver-manager openpyxl pillow`
