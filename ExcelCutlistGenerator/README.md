# ExcelCutlistGenerator
A script to run along side an excel template that allows the user to enter part numbers, lengths, and quantities and then nest them according to predefined stock lengths

## Features Wish List
* Check for duplicate part numbers and alert the user
* Include cut angles in nesting considerations
* Output saw programming data

# Cut Angle Development

## Orientation In Saw
When considering the end angles of the part in the nest there the side of the stock that the cut is on needs to be considered.
The naming convention that Advance Steel uses is Web Side or Flange Side. The side of the stock that the cut angle is on doesn't
need to be considered for stock that is symmetric (ex. square tube or equal leg angle). But for non-symmetric stock it does.
<br>
<br>
To handle this a cut orientation will be added to the part object that will represent the orientation of the stock in the saw.
When nesting the parts, each stick will be assigned a cut orientation and only parts with that orientation can be considered for that stick.
<br>
<br> 
### **0 = orientation doesn't matter**
This can be used for either parts with no angle cuts or parts being cut from symmetric stock.  These parts
can be assigned to a stick with a 1 or 2 orientation. 
<br>
<br>

### **1 = width/flange side down in the saw**
This would either be the flange of the beam or the short side of un-symmetric stock being horizontal in the saw
<br>
\***<br>
\* *<br>
\* *<br>
\* *<br>
\* *<br>
\* * * * * *<br>
\* * * * * *<br>
___________________<br>

<br>

### 2 = length/web side down in saw
This would either be the web side of a beam or the tall side of un-symmetric stock being horizontal in the saw

\***<br>
\* *<br>
\* *<br>
\* * * * * * * * * *<br>
\* * * * * * * * * *<br>
___________________<br>

<br> 
<br>

### Parts with cuts on different planes
An example of this would be a part that on one end has a cut on the web side and on the other has a cut on the flange side.
Cutting this can't be automated because the stock would need to be flipped between the first and second cut. 
Below is how these parts will be handled.

* They can be assigned to a stick with either orientation.
* The will be added to the end of their respective stick.
* They will have a warning placed next to them in the cut list.
* They will not be included in the saw code.


## Interfacing with Advance Steel Cutlist
Advance Steel cutlist assign the cut angles in 4 columns Left Web, Right Web, Left Flange, and Right Flange.
It will be necessary to interpret what those values mean for different stock.

A non-zero value in one or more flange and one or more web column indicates that the first and second cut are on separate planes.

### Square Tube
For square tube, the values for the web columns and flange columns can be used interchangeably without any modification.
This is ofcourse assuming that there are only values in either the flange columns or web columns. 

### Rectangle Tube

### Equal Leg Angle
For equal leg angle the angle values can be in either the web or flange columns.  If the values are in the 
web columns or flange columns is determined by the detailing orientation. Although more testing needs to be done,
it seems that if the cut feature has the angle in Y, the values are as if the bottom leg is sitting on the saw and pointed at you (vertical leg against the back jaw of vice).
If the cut feature angle is in the Z axis, it is the opposite. 






