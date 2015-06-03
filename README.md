# CalcMaster
SharePoint Calculated Column Editor by ViewMaster365.com

**This code only works on existing Calculated Columns, it will not work if you have just created a new column**

##Use

Activated on a SharePoint FldEdit.aspx page this code will attach itself to the Formula textarea.
On every keyup event it will (try to) save the Formula and provide immediate feedback.

The OK button is no longer needed, because a correct Formula is always saved.

##Installation

###Save the Javascript file to your own server
* Download the VM365_CalcMaster.js script file to a location within your SharePoint/Office365 environment
* Copy the full URL to this file (e.g: //vm365.sharepoint.com/Site pages/VM365_CalcMaster.js)

###Trigger with a Browser Bookmarklet (manual activation)
* Right click in the Favourites/Bookmark bar to Add a New Page (on IE add the WebPage as Favorite then overwrite the address with the JavaScript below)
* Name it CalcMaster
* fill the URL with javascript code (don't forget to change the URL):

    javascript:(function(){var url='YOUR_URL_HERE',jsCode=document.createElement('script');jsCode.setAttribute('src', url);document.body.appendChild(jsCode);}())
    
###Trigger with Chrome Tampermonkey (auto activation)
Install the [https://tampermonkey.net/](Tampermonkey plugin) (alas not possible in IE) and create a script:

    // ==UserScript==
    // @name        CalcMaster
    // @author      Danny Engelman  
    // @namespace   http://ViewMaster365.com/
    // @version     0.1
    // @description Calculated Column Editor
    // @match       *FldEdit.aspx*
    // @copyright   2012+, You
    // @require     http://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js
    // @require     https://gist.github.com/raw/2625891/waitForKeyElements.js
    // @grant       GM_addStyle
    // ==/UserScript==
    waitForKeyElements ( '#onetidIODefTextValue1' , function(){
        //attach bookmarklet code to page and trigger execute
        var VM365_CalcMasterDirectory='https://365csi.sharepoint.com/sites/VM/SitePages/CalcMaster/';

        var jsCode = document.createElement('script'); 
        jsCode.setAttribute('src', VM365_CalcMasterDirectory + 'VM365_CalcMaster.js'  );
        document.body.appendChild(jsCode);
    } );


