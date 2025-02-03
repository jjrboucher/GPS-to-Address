This script will launch a GUI where you point to an Excel file with a single worksheet in it.

The worksheet must have a column with latitude and one for longitude. The menu will present two pull down menus where you can pick the column for lat and the one for long.

It accepts any of the following three formats (with lat/long in different columns, no "/" between the two):<br>
46.203361 / 17.341428<br>
46 ° 12' 12.10" N / 17 ° 20' 29.14"<br>
46 deg 12' 12.10" N / 17 deg 20' 29.14"<br>
<br>
You execute the script. It will go through and convert each to an address. If it can't convert a value, it will note that. The results will be in a new column called Address.
