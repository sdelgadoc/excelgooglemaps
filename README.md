# excelgooglemaps

Simple Excel functions written in VBA to call the [Google Maps API](https://developers.google.com/maps/documentation) easily.
* Returns the driving distance between two addresses
* Cleans up addresses to a Google Maps compliant format
* Returns the postal code of an address

## Setup
First, clone this repository onto your system.

Then, in an Excel spreadsheet, type Alt+F11 to open the VBA development environment.

In the VBA development environment, click on File > Import File... and navigate to the location of the `DistanceFunctions.bas` file from this repo.



Close the VBA development environment, and you will be able to to use the Google Maps functions in your spreadsheet.

## Usage
### GetDistance
Function that returns the driving distance in meters between two addresses
#### Parameters
* Source address: The Google Maps compliant address of the source location
* Destination address: The Google Maps compliant address of the destination location
### GetAddress
Function returns a Google Maps compliant address based on minimal information
#### Parameters
* Address fragment: String of minimal address information
### GetPostalCode
Function returns the zip code of a Google Maps compliant address
#### Parameters
* Address: A Google Maps compliant address string

## Maintainer
Santiago Delgado  ([@santiagodc](https://twitter.com/santiagodc))

## License
MIT

## Disclaimer
This repo has no affiliation with Microsoft or Google
