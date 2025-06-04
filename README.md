# MMS WebScrapper

These are a set of Python scripts to scrape several charts and data from MMS for the exam board reports. This is part of the role of the Exam Officer.

## Requirements for the code to run:

1. You need the list of modules in a format like: "GG4258" no spaces
2. The link for charts is https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG3214/Final+grade/GraphPage, and the link for the grade table is https://mms.st-andrews.ac.uk/mms/module/2024_5/S2/GG1002/Final+grade/ you will notice both links have the module code that can act as parameter.
3. There are two Python scripts that web scraped the info required for the Excel that is shared with the externals, although there is still some manual work to do, it saves tons of time.
4. module_summary_scraper.py does the extraction of the statistics from the module Final grade page.
module_charts_downloader.py does the extration of both charts provided by MMS, created a folder to stores those and then create a excel file with individual sheets for each module and both charts.

The manual work implies joining both Excel files and also adjusting the information from the summary scraper. As there are modules with different structures, the table can be completely different in some modules. Therefore, I get all the info and then manually validate what the correct info is for the unusual modules.

You can adapt the code to update the list of modules per semester and also the AY, as for now, the link is consistent across modules.

This still requieres more work and some script adaptation but it should save a lot of time with the reporting.
