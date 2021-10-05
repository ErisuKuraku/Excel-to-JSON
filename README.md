# Excel to JSON file

This was done for a college task, but since there wasn't much info for those having issues. Here is my spreadsheet.

Feel free to look and compare, this project is under MIT licence.

# System Requirements

- Confirmed working on Excel 2016 +
- Only Tested on Windows, if it works on Mac please let me know
- Developer Tab Enabled >> [See Here](https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)

# Special Thanks!

Special thanks to [VBA-tools](https://github.com/VBA-tools/VBA-JSON) for their great module to Excel VBA, This should be applied by default but if not then the .bs file can be found on their repo.

Also to [Excelerator Soulutions](https://excelerator.solutions/) for their guide

# Possible Issues

Here are the issues we encountered during development:

|     **Error**           |     **Reason**                                                      |     **Solution**                                                                  |
|-------------------------|---------------------------------------------------------------------|-----------------------------------------------------------------------------------|
|     Permission          |     For security reasons,   Excel can’t access critical folders.    |     Change the   Directory to somewhere less guarded (e.g. documents, desktop)    |
|     Invalid   Object    |     A command   that Excel has called doesn’t exist                 |     Make sure the   JSON Converter file is imported under modules                 |
|                         |                                                                     |                                                                                   |

If you encounter anything wrong with my code then post an issue
