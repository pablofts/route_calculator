# ROUTE CALCULATOR vba -traveling salesman problem-
Calculates the shortest route that goes through a list of places. The input is an array containing the coordinates of each place the route should go by. The output is a matrix containing the distances and times between places and a link to display the route in google maps.

# You will need to
* Enable developer tab in Excel. Instructions at https://msdn.microsoft.com/en-us/library/bb608625.aspx.
* Add "Microsoft XML, v6.0" as a Reference. *Tools* -> *References*.
* Have a credit card.
* Go to google console and get a Google Maps API DIstance Matrix at https://developers.google.com/maps/documentation/distance-matrix/start (it should remain free unless you use it a lot -i run it with around twenty places really offten and it has remain charge free-)

# Installation
* Import the .bas files into your project.

# Usage
* a_input module: within the vba text editor, you will need to provide an array with the places list in the, where asked to.
* b_timedistance_matrix: provide the API key, the traveling mode and region in the following variables: apikey, modo, region.
* Run the code.
* The output is an array, so you may want to print it in a worksheet.

# Credits to:
* Mathew Moran, whose code I used to request the times and distances to google maps. Public code at https://pulseinfomatics.com/new-use-vba-to-retrieve-distances-between-multiple-addresses-in-excel/
