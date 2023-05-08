# energyCalc

MS Excel Add-In with functions for calculations related to energy consumption

## Funcions

### VERBRAUCHSPROGN()

The function can be used to calculate the probable consumption of energy (electricity, gas or district heating) based on a consumption value for a partial period of a year. The underlying weighting tables for electricity and gas are based on smoothed own measurement results over almost two decades. However, you can also use your own weight values as a basis.
The start date must be less than the end date. The period (start date to prognose Date) can be a maximum of one year. Corresponding to the fact that energy bills are often distributed on a rolling basis over the year (starting from the date of the start of delivery), this function can also calculate across years. The function calculates on a daily basis. Leap years are taken into account.

![function arguments VERBRAUCHSPROGN](/assets/function_arguments_3.jpg)

Two optional input parameters are available:

* *Sparte*: The weighting table for electricity (own values) is stored as the default basis for the forecast. With the value "2" the weighting table for gas (own values), with the value "3" for district heating (degree day numbers).
* *IndivWerte*: You can use it to transfer twelve individual weighting numbers (January to December) and overwrite the default values (percent in decimal fractions). You can use your own tables or on of the altenatives below.

#### Examples

![examples](/assets/examples_3.jpg)

#### Weight tables

##### Electricity

|Month|build-in|SAP IS-U(?)|[Musterhaushalt](https://www.musterhaushalt.de/durchschnitt/stromverbrauch/)|[NEW](https://www.new.de/fileadmin/user_upload/new.de/Dokumente/Service/Mengenaufteilung_fuer_Stromkunden_bei_Preisaenderungen_x.pdf)|[ViVi](https://www.vivi-power.de/documents/vivi-abgrenzungstrom2013.pdf)|
|-|-|-|-|-|-|
|Januar   |0.0939|0.0994|0.0932|0.1030|0.0950|
|Februar  |0.0833|0.0931|0.0857|0.0890|0.0900|
|März     |0.0881|0.0843|0.0892|0.0920|0.0850|
|April    |0.0820|0.0742|0.0806|0.0830|0.0800|
|Mai      |0.0803|0.0739|0.0783|0.0780|0.0800|
|Juni     |0.0739|0.0725|0.0756|0.0700|0.0750|
|Juli     |0.0713|0.0729|0.0777|0.0690|0.0750|
|August   |0.0755|0.0739|0.0761|0.0710|0.0750|
|September|0.0803|0.0742|0.0792|0.0730|0.0800|
|Oktober  |0.0852|0.0843|0.0852|0.0840|0.0850|
|November |0.0895|0.0931|0.0888|0.0870|0.0850|
|Dezember |0.0967|0.1042|0.0904|0.1010|0.0950|

##### Gas

|Month|build-in|SAP IS-U(?)|[Haustechnikdialog](https://www.haustechnikdialog.de/Forum/p/728559#p728559)|[Statista](https://de.statista.com/statistik/daten/studie/160067/umfrage/verbrauch-von-heizenergie-nach-monaten/)|[ViVi](https://www.vivi-power.de/documents/vivi-abgrenzunggas2013.pdf)|
|-|-|-|-|-|-|
|Februar  |0.1643|0.1314|0.1810|0.1610|0.1800|
|März     |0.1125|0.1253|0.1540|0.1300|0.1600|
|Januar   |0.1984|0.1184|0.1400|0.1250|0.1300|
|April    |0.0451|0.0921|0.0900|0.0810|0.0700|
|Mai      |0.0208|0.0641|0.0370|0.0350|0.0500|
|Juni     |0.0079|0.0392|0.0000|0.0220|0.0200|
|Juli     |0.0058|0.0214|0.0000|0.0170|0.0200|
|August   |0.0059|0.0202|0.0000|0.0160|0.0200|
|September|0.0164|0.0450|0.0160|0.0520|0.0200|
|Oktober  |0.0762|0.0859|0.0870|0.0840|0.0700|
|November |0.1541|0.1174|0.1280|0.1220|0.1100|
|Dezember |0.1962|0.1396|0.1670|0.1550|0.1500|

##### District heating

|Month|build-in ([Gradtagszahlen](#gradtagsz))|
|-|-|
|Februar  |0.1700|
|März     |0.1500|
|Januar   |0.1300|
|April    |0.0800|
|Mai      |0.0400|
|Juni     |0.0130|
|Juli     |0.0135|
|August   |0.0135|
|September|0.0300|
|Oktober  |0.0800|
|November |0.1200|
|Dezember |0.1600|


### GRADTAGSZ()

Degree day figures represent the connection between the room and the outside air temperature and serve as an aid for determining the heating costs in the event of a change of user during a specific billing period. If there is a change of tenant within the billing period and the interim reading is not possible, for example, degree day figures are a suitable means of dividing the basic costs of heat consumption fairly and in accordance with the rules.

![figures](/assets/table.jpg)

This function calculates the sum of the number of daily temperature figures (Gradtagszahlen) between two dates.

The start date must be less than the end date. The period can be a maximum of one year. A calculation usually only takes place within a calendar year, this function can also calculate across years. The function calculates on a daily basis. Leap years are taken into account.

![function arguments GRADTAGSZ](/assets/function_arguments_2.jpg)

Two optional input parameters are available:

* *Quotient* (default false): if you set this to "1" (true), the fraction will be displayed instead of the integer
* *IndivWerte*: You can use it to transfer twelve of your own degree day numbers (January to December) and overwrite the default values

#### Examples

![examples](/assets/examples.jpg)

### GASABR()

Calculates the calorific value of measured cubic meters of gas. The calorific value describes the amount of energy that is released during the combustion of gas and the subsequent cooling of the exhaust gases (heat of condensation). The gas condensing value is given in kilowatt hours per cubic meter (kWh/m³). Beside the calorific value of the gas the conversion factor (ZustandsZahl) describes the normalization of the gas for height and temperature.

![function arguments GASABR](/assets/function_arguments_1.jpg)

In addition to the calculation, this function provides the plausibility check of the values for the calorific value and the conversion factor:

* calorific value must be > 8 and < 14 kWh/m³
* conversion factor depending on:
  * air pressure between sea level and 1000 m (1,013.25 and 898.76 hPa)
  * temperature between -30° and +40° centigrade

The function result is the calorific value of the normalized volume in kWh

## License

This script is released under MIT license.

## Author

Copyright© 2023 Volker Püschel
