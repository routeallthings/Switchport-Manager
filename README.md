# Switchport Manager

The goal of this script was to create an easy to use switchport manager to mass assign VLANs based on vendor, find where VLANs and ports were out of alignment, and create reporting on this.

## Getting Started

Look at the template XLSX. Fill in the data and populate the Device Config page.

Step 1. Fill in the XLSX data
Step 2. Run the script
Step 3. Drink a beverage of your choosing while laughing like a maniac

Report any issues to my email and I will get them fixed.

### Prerequisites

GIT (This is required to download the XLHELPER module using a fork that  I made for compatibility with Python 2.7)
XLHELPER
OPENPYXL

## Deployment

Just execute the script and answer the questions

## Features
- XLSX-based import
- Export to XSLX for reporting
- Mass change ports based on vendor mac
- Find issues where vendor mac and port are not aligned

## *Caveats
- None

## Versioning

VERSION 1.0
Currently Implemented Features
- XLSX-based import
- Export to XSLX for reporting
- Mass change ports based on vendor mac
- Find issues where vendor mac and port are not aligned

## Authors

* **Matt Cross** - [RouteAllThings](https://github.com/routeallthings)

See also the list of [contributors](https://github.com/routeallthings/Switchport-Manager/contributors) who participated in this project.

## License

This project is licensed under the GNU - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Thanks to HBS for giving me a reason to write this.
