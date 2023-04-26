# energyCalc

MS Excel Add-In with functions for calculations related to energy consumption

## Functions

Only one at the moment but more will be added...

### GRADTAGSZ()

Degree day figures represent the connection between the room and the outside air temperature and serve as an aid for determining the heating costs in the event of a change of user during a specific billing period. If there is a change of tenant within the billing period and the interim reading is not possible, for example, degree day figures are a suitable means of dividing the basic costs of heat consumption fairly and in accordance with the rules.

![figures](/assets/table.jpg)

This function calculates the sum of the number of daily temperature figures (Gradtagszahlen) between two dates.

The start date must be less than the end date. The period can be a maximum of one year. A calculation usually only takes place within a calendar year, this function can also calculate across years. The function calculates on a daily basis. Leap years are taken into account.

![function arguments](/assets/function_arguments.jpg)

Two optional input parameters are available:

* *Quotient* (default false): if you set this to "1" (true), the fraction will be displayed instead of the integer
* *DiffValues*: You can use it to transfer twelve of your own degree day numbers (January to December) and overwrite the default values

![examples](/assets/examples.jpg)

## License

This script is released under MIT license.

## Author

Copyright© 2023 Volker Püschel
