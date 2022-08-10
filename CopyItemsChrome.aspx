<div id='myModal' class='modal s4-wpcell'> 
	<div class='modal-content'> 
		<span class='modal-close'>&times;</span> 
		<h2>Copy items to Outbound-TM</h2>
		<p>Please select the destination document set:</p>
		<div class="modal-center">
		<select name='DocumentSets' style="width:40%;" id='DocumentSets'></select> 
		<button type="button" id='CopyButton' style="margin-left:20px;" onclick="copyFiles()">Start Copy</button>
		</div>
		<span id='dvShowMessage' class='modalMessages' ></span> 
		<span id='modalFileExists' class='modalMessagesError'>The following file(s) already exists in the destination, they have not been transferred:</span>
		<span id='modalCopyError' class='modalMessagesError'>There was an issue copying, please try again.</span>
		<span id='modalFileExistList' class='modalMessagesError'></span>
		
	</div> 
</div>


<script type="text/javascript" src="/_layouts/15/sp.js"></script>
<script type="text/javascript" src="/_layouts/15/SP.RequestExecutor.js"></script>

  
<style>
	.modal {
	  display: none; /* Hidden by default */
	  position: fixed; /* Stay in place */
	  z-index: 100; /* Sit on top */
	  padding-top: 100px; /* Location of the box */
	  left: 0;
	  top: 0;
	  width: 100%; /* Full width */
	  height: 100%; /* Full height */
	  overflow: auto; /* Enable scroll if needed */
	  background-color: rgb(0,0,0); /* Fallback color */
	  background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
	}

	.modal-center {
	  position: relative;
	  left: 20%;
	}

	/* Modal Content */
	.modal-content {
	  background-color: #fefefe;
	  margin: auto;
	  padding: 20px;
	  border: 1px solid #888;
	  width: 60%;
	}

	/* The Close Button */
	.modal-close {
	  color: #aaaaaa;
	  float: right;
	  font-size: 28px;
	  font-weight: bold;
	}

	.modal-close:hover,
	.modal-close:focus {
	  color: #000;
	  text-decoration: none;
	  cursor: pointer;
	}

	.modalMessages {
		display: block;
		text-align: center;
		margin-top: 15px;
		font-size: 13px;
	}
	
	.modalMessagesError {
		display: block;
		text-align: center;
		margin-top: 15px;
		font-size: 13px;
		color: red;
	}

	#modalFileExists {
		display: none;
	}

	#modalCopyError {
		display: none;
	}
</style>
 
<script type="text/javascript">

/*******************************************/
// ASSUMED /SiteAssets/jquery.min.js is in the SP Web 
/*******************************************/
targetLibraryUrl = "Target"; //Change to the target Library
targetLibraryName = "Target";
/*******************************************/

// Get the modal
var modal = document.getElementById("myModal");

// Get the <span> element that closes the modal
var span = document.getElementsByClassName("modal-close")[0];

// When the user clicks on <span> (x), close the modal
span.onclick = function() {
	document.getElementById('dvShowMessage').innerHTML = "";
	modal.style.display = "none";
	modalFileExists.style.display = "none";
	modalCopyError.style.display = "none";
	document.getElementById('modalFileExistList').innerHTML = "";	
}

// When the user clicks anywhere outside of the modal, close it
window.onclick = function(event) {
  if (event.target == modal) {
	document.getElementById('dvShowMessage').innerHTML = "";
	modal.style.display = "none";
	modalFileExists.style.display = "none";
	modalCopyError.style.display = "none";
	document.getElementById('modalFileExistList').innerHTML = "";
  }
}

//Wait for sp.ribbon.js to be loaded
ExecuteOrDelayUntilScriptLoaded(function () { 
	var tag = document.createElement("script");
	tag.src = _spPageContextInfo.webAbsoluteUrl + "/" + "/SiteAssets/jquery.min.js" ;
	document.getElementsByTagName("head")[0].appendChild(tag);

	var pm = SP.Ribbon.PageManager.get_instance();
	pm.add_ribbonInited(function() { //Adds a handler to the handle the RibbonInited event.
	var ribbon = (SP.Ribbon.PageManager.get_instance()).get_ribbon();
		createTab(ribbon);
	});
	
	// Populate the document sets in the Copy Items drop down menu
	populateDocumentSets();
		
}, 'sp.ribbon.js');

function createTab(ribbon) {
    var tabId = "CopyFiles.Tab.Id";
    var tabGroupId = "CopyFiles.Tab.Group.Id";
    var tabText = "Copy Files";
    var tabGroupText = "Copy Files";
    var tabButtonText = "Copy Files";

    var tab = new CUI.Tab(ribbon, tabId, tabText, "Copy Files Tab", 'Tab.Attach.Command', false, '', null, null);
    ribbon.addChild(tab);

    var group = new CUI.Group(ribbon, tabGroupId, tabGroupText, "Group Copy Files", 'Tab.Attach.Group.Command', null);
    tab.addChild(group);

    var layout = new CUI.Layout(ribbon, 'Tab.AttachNews.Layout', 'The Layout');
    group.addChild(layout);

    var section = new CUI.Section(ribbon, 'Tab.AttachNews.Section', 2, 'Top'); //2==OneRow
    layout.addChild(section);
    var controlProperties = new CUI.ControlProperties();
    controlProperties.Command = 'Tab.AttachNews.Button.Command';
    controlProperties.Id = 'Tab.AttachNews.ControlProperties';
    controlProperties.TemplateAlias = 'o1';
    controlProperties.ToolTipDescription = tabButtonText;
    controlProperties.Image32by32 = '_layouts/15/images/centraladmin_applicationmanagement_serviceapplications_32x32.png';
    controlProperties.ToolTipTitle = tabButtonText;
    controlProperties.LabelText = tabButtonText;
    var button = new CUI.Controls.Button(ribbon, 'Tab.AttachNews.Button', controlProperties);
    button.set_enabled(true);
    var controlComponent = button.createComponentForDisplayMode('Large');
    var row1 = section.getRow(1);

    row1.addChild(controlComponent);
    group.selectLayout('The Layout');

    SelectRibbonTab(tabId, true); //Mark new tab as selected
    controlComponent.set_enabled(true);
    var btn = button.getDOMElementForDisplayMode("Large");
	
	//attach click event to the newly created ribbon button
    $(btn).click(function () { 
		document.getElementById("CopyButton").disabled = false;
		var context = SP.ClientContext.get_current();
		var selectedItemIds = SP.ListOperation.Selection.getSelectedItems(context); 
		if(selectedItemIds.length != 0){
			var modal = document.getElementById("myModal");
			modal.style.display = "block";
		}
		else{
			alert("Please select at least one item to copy.");
			return false;
		}
		
	}); 
}

function populateDocumentSets() {
    var context = new SP.ClientContext.get_current();
    var oList = context.get_web().get_lists().getByTitle(targetLibraryName);
        
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query><Where><Eq><FieldRef Name=\'ProgId\'/>' + 
        '<Value Type=\'Text\'>Sharepoint.DocumentSet</Value></Eq></Where></Query>' + 
        '</View>'
    );
    this.collListItem = oList.getItems(camlQuery);
        
    context.load(collListItem);
    context.executeQueryAsync(
        Function.createDelegate(this, this.getDocumentSetsQuerySucceeded), 
        Function.createDelegate(this, this.onQueryFailed)
    ); 
}

function getDocumentSetsQuerySucceeded(sender, args) {
    var listItemInfo = '';
    var listItemEnumerator = collListItem.getEnumerator();
	var dropDown = document.getElementById("DocumentSets");
	
    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
		var option = document.createElement('option');
		option.text = option.value = oListItem.get_item('Title');
		dropDown.add(option, 0);
    }
}

function onQueryFailed(sender, args) {
    alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

async function copyFiles() 
{
	var DocumentSets = document.getElementById("DocumentSets");
	var DocumentSet = DocumentSets.options[DocumentSets.selectedIndex].text;
	var context = SP.ClientContext.get_current();
	var selectedItemIds = SP.ListOperation.Selection.getSelectedItems(context); 
	var DestinationUrl = _spPageContextInfo.webAbsoluteUrl + "/" + targetLibraryUrl + "/" + DocumentSet + "/";

	var web = context.get_web();
	var transferCount = 0;
	var countFilesToMove = 0;

	document.getElementById('modalFileExists').style.display = 'none';
	document.getElementById('modalCopyError').style.display = 'none';
	document.getElementById('dvShowMessage').innerHTML = '   Copying...';
	document.getElementById("CopyButton").disabled = true;

	for (idx of selectedItemIds)
	{
		countFilesToMove += 1;
		var source = $("#" + idx.id + " a")[0].href.toString();
		var target = DestinationUrl + $("#" + idx.id + " a")[0].href.split("/")[$("#" + idx.id + " a")[0].href.split("/").length - 1].toString(); 
		var CopyFileName = source.split("/").pop();
		
		targetUrl = "/" + target.split(".com/")[1]
		var ofile = web.getFileByServerRelativeUrl(targetUrl);
		context.load(ofile);
		var FileFound = "";
		
		await new Promise((resolve, reject) => {
			context.executeQueryAsync((x) => {
				FileFound = "found";
				resolve();
			}, (error) => {
				console.log(error);
				resolve();
			});
			});
		
		if(FileFound == "found"){
			//Populate the File Exists list on the Copy Window
			document.getElementById("modalFileExists").style.display = "block";
			document.getElementById('modalFileExistList').innerHTML += CopyFileName + "</br>";
		}
		else{
			//Add the file to the copy util
			SP.MoveCopyUtil.copyFile(context, source, target, true );
			transferCount += 1;
		}
	}
	
	if(transferCount > 0){
		context.executeQueryAsync(
			function(){
				document.getElementById('dvShowMessage').innerHTML = "<font color ='green'>Transfer complete</font>";
			}, 
			function(sender, error){
				var message = error.get_message();
				if(message.includes("file already exists")){
					document.getElementById("modalFileExists").style.display = "block";
					document.getElementById('dvShowMessage').innerHTML = "";
					var requestBody = sender.get_webRequest()._body;
					var bodyXml = $.parseXML(requestBody)
					$xml = $(bodyXml)
					var sourceFileName = $xml.find("Parameter")[0].innerHTML.split("/").pop()
					document.getElementById('modalFileExistList').innerHTML += sourceFileName + "</br>";
				}
				else{
					document.getElementById("modalCopyError").style.display = "block";
					document.getElementById('dvShowMessage').innerHTML = "";
				}
			}
		)	
	}
	else{
		document.getElementById('dvShowMessage').innerHTML = "";
	}
}

</script>