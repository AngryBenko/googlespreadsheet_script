# googlespreadsheet_script
A simple script I wrote to automate some formatting on a specific schedule based on certain conditions:

* Event types (or under the Stock # column) were as followed: a stock number, BFP, or RETURN. Now each of these also required specific formatting conditions as well
  * A stock number is a green background and an array of checkboxes
  * A BFP is a pink background with no array of checkboxes
  * A RETURN is a black background with no array of checkboxes
  ![checkboxes](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/checkboxes.png)
* Any events that occured in previous dates will be greyscaled
* Events that occur of different dates will be seperated with a thick border
* Different dates have color coordinated using the script proprties in Google Spreadsheet
![scriptprop](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/script_properties.png)

This schedule started out very simple with minimal amount of workers and events. As the team and company grew, so did the amount of scheduled events. This prompted the need to rework the current schedule to fit the new workload. 

![before](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/before_schedule.png)

This was the schedule that the team decided on that would fit the new workload. You can see the that different days have different colors and are seperated with a border. Stock numbers which are deliveries are colored as green, BFP are pink, and RETURNS are black. Previous dates are greyscaled. 

![wip](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/wip_schedule.png)
![current](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/current_schedule.png)

All of these formatting changes can be applied using Google scripts onEdit() function. However, when there are many events filled into one spreadsheet, this can cause the script to run slow. To prevent this, I added a custom menu option that you can click to apply the changes after you are done editing the information on the spreadsheet.

![menu](https://github.com/AngryBenko/googlespreadsheet_script/blob/master/images/menu.png)
