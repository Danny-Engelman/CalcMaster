# CalcMaster - Calculated Column Editor

Activated on a SharePoint FldEdit.aspx page this code will attach itself to the Formula textarea.  

##### features
* On every keyup event it will (try to) save the Formula and provide immediate feedback.
* Paste plain JavaScript code in an empty formula field and wrap in Formula strings with the action button.  
*Why I added this?* See [www.viewmaster365.com/#/How](http://www.viewmaster365.com/#/How)
* The OK button is no longer needed, because a correct Formula is always saved.

*This code only works on existing Calculated Columns,  
it will not work if you have just created a new column, save it first*

![](http://i.imgur.com/RPvRrDr.jpg)


###Version History
* june  3 2015 - extracted primary code from my personal advanced Calcmaster
* june  5 2015 - got rid of all jQuery dependencies
* june 18 2015 - added console, added link to [SharePoint Calculated Column Functions Syntax List](http://viewmaster365.com/365coach/#/Calculated_Column_Functions_List)

##Installation
Since this code executes within your browser I have not included the Bookmarklet. Built it your self so you understand the security implications:

* Copy the code from [VM365_CalcMaster.js](https://raw.githubusercontent.com/Danny-Engelman/CalcMaster/master/VM365_CalcMaster.js)
* Paste it at [the Bookmarkleter](http://chriszarate.github.io/bookmarkleter/) (*use default settings, no need to include jQuery*)

##OR
###Save the Javascript file to your own server
* Download the [VM365_CalcMaster.js](https://raw.githubusercontent.com/Danny-Engelman/CalcMaster/master/VM365_CalcMaster.js) script file to a Library within your SharePoint/Office365 environment
* Copy the full URL to this file (e.g: //vm365.sharepoint.com/Pages/VM365_CalcMaster.js)

###Trigger with a Browser Bookmarklet (manual activation)
* Right click in the Favourites/Bookmark bar to Add a New Page (on IE add the WebPage as Favorite then overwrite the address with the JavaScript below)
* Name it CalcMaster (or anything you want)
* fill the URL with javascript code (don't forget to change the URL):

    javascript:(function(){var url='YOUR_URL_HERE',jsCode=document.createElement('script');jsCode.setAttribute('src', url);document.body.appendChild(jsCode);}())
    
###Trigger with Chrome Tampermonkey (auto activation)
Tampermonkey can watch what you are doing in the browser and trigger code when you browse to FldEdit.aspx page

* Install the [Tampermonkey plugin](https://tampermonkey.net/) (alas not possible in IE) 
* Add the [Tampermonkey script](https://github.com/Danny-Engelman/CalcMaster/blob/master/Tampermonkey)  do not forget to change yout URL
