# SharePoint-Javacript-CopyItems-Solution
This solution adds a "Copy Files" button to the ribbon, to allow a user to copy items to a target library in SharePoint.
![Example](https://raw.githubusercontent.com/martinpyman/SharePoint-Javacript-CopyItems-Solution/master/screen1.png | width=100)

A drop down is populated with the Document Sets found in the Target library.
![Drop down](https://raw.githubusercontent.com/martinpyman/SharePoint-Javacript-CopyItems-Solution/master/screen2.png | width=110)

This solution is useful for SharePoint 2016 which is lacking this functionality and you are restricted in deploying a SharePoint app.
(This has only been tested on SP2016!)

First:
Update the"targetLibraryUrl" and the "targetLibraryName" in the CopyItems.aspx (for internet explorer users) or CopyItemsChrome.aspx to the Name and Url of the library you wish items to be migrated into. 

To deploy:
- Copy the items to the /SiteAssets library
- Navigate to the source library where you wish to deploy
- Edit the view you wish the extension to be added to, add a content editor webpart to the bottom of the view
- In the Content Editor webpart link to the aspx files CopyItems.aspx (this is supported with IE)
If you are using a modern browser use the "CopyItemsChrome.aspx"  !! 

Save the page and select an item, the Copy Items button should instantly appear in the ribbon.

