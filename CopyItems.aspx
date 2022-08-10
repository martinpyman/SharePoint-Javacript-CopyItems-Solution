<div id='myModal' class='modal'> 
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


<script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
<script type="text/javascript" src="/_layouts/15/sp.js"></script>
<script type="text/javascript" src="/_layouts/15/SP.RequestExecutor.js"></script>
<script type="text/javascript" src="/teams/projectsite1/SiteAssets/jquery.min.js"></script>
<script type="text/javascript" src="/teams/projectsite1/SiteAssets/runtime.js"></script>
<script type="text/javascript" src="/teams/projectsite1/SiteAssets/polyfill.min.js"></script>

  
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


function _createForOfIteratorHelper(o, allowArrayLike) { var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"]; if (!it) { if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") { if (it) o = it; var i = 0; var F = function F() {}; return { s: F, n: function n() { if (i >= o.length) return { done: true }; return { done: false, value: o[i++] }; }, e: function e(_e) { throw _e; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var normalCompletion = true, didErr = false, err; return { s: function s() { it = it.call(o); }, n: function n() { var step = it.next(); normalCompletion = step.done; return step; }, e: function e(_e2) { didErr = true; err = _e2; }, f: function f() { try { if (!normalCompletion && it.return != null) it.return(); } finally { if (didErr) throw err; } } }; }

function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }

function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) { arr2[i] = arr[i]; } return arr2; }

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

/*******************************************/
// ASSUMED /SiteAssets/jquery.min.js is in the SP Web 

/*******************************************/
targetLibraryUrl = "Target"; //Change to the target Library

targetLibraryName = "Target";
/*******************************************/
// Get the modal

var modal = document.getElementById("myModal"); // Get the <span> element that closes the modal

var span = document.getElementsByClassName("modal-close")[0]; // When the user clicks on <span> (x), close the modal

span.onclick = function () {
  document.getElementById('dvShowMessage').innerHTML = "";
  modal.style.display = "none";
  modalFileExists.style.display = "none";
  modalCopyError.style.display = "none";
  document.getElementById('modalFileExistList').innerHTML = "";
}; // When the user clicks anywhere outside of the modal, close it


window.onclick = function (event) {
  if (event.target == modal) {
    document.getElementById('dvShowMessage').innerHTML = "";
    modal.style.display = "none";
    modalFileExists.style.display = "none";
    modalCopyError.style.display = "none";
    document.getElementById('modalFileExistList').innerHTML = "";
  }
}; //Wait for sp.ribbon.js to be loaded


ExecuteOrDelayUntilScriptLoaded(function () {
  var pm = SP.Ribbon.PageManager.get_instance();
  pm.add_ribbonInited(function () {
    //Adds a handler to the handle the RibbonInited event.
    var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon();
    createTab(ribbon);
  }); // Populate the document sets in the Copy Items drop down menu

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
  var btn = button.getDOMElementForDisplayMode("Large"); //attach click event to the newly created ribbon button

  $(btn).click(function () {
    document.getElementById("CopyButton").disabled = false;
    var context = SP.ClientContext.get_current();
    var selectedItemIds = SP.ListOperation.Selection.getSelectedItems(context);

    if (selectedItemIds.length != 0) {
      var modal = document.getElementById("myModal");
      modal.style.display = "block";
    } else {
      alert("Please select at least one item to copy.");
      return false;
    }
  });
}

function populateDocumentSets() {
  var context = new SP.ClientContext.get_current();
  var oList = context.get_web().get_lists().getByTitle(targetLibraryName);
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ProgId\'/>' + '<Value Type=\'Text\'>Sharepoint.DocumentSet</Value></Eq></Where></Query>' + '</View>');
  this.collListItem = oList.getItems(camlQuery);
  context.load(collListItem);
  context.executeQueryAsync(Function.createDelegate(this, this.getDocumentSetsQuerySucceeded), Function.createDelegate(this, this.onQueryFailed));
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

function copyFiles() {
  return _copyFiles.apply(this, arguments);
}

function _copyFiles() {
  _copyFiles = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee() {
    var DocumentSets, DocumentSet, context, selectedItemIds, DestinationUrl, web, transferCount, countFilesToMove, _iterator, _step, source, target, CopyFileName, ofile, FileFound;

    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            DocumentSets = document.getElementById("DocumentSets");
            DocumentSet = DocumentSets.options[DocumentSets.selectedIndex].text;
            context = SP.ClientContext.get_current();
            selectedItemIds = SP.ListOperation.Selection.getSelectedItems(context);
            DestinationUrl = _spPageContextInfo.webAbsoluteUrl + "/" + targetLibraryUrl + "/" + DocumentSet + "/";
            web = context.get_web();
            transferCount = 0;
            countFilesToMove = 0;
            document.getElementById('modalFileExists').style.display = 'none';
            document.getElementById('modalCopyError').style.display = 'none';
            document.getElementById('dvShowMessage').innerHTML = '   Copying...';
            document.getElementById("CopyButton").disabled = true;
            _iterator = _createForOfIteratorHelper(selectedItemIds);
            _context.prev = 13;

            _iterator.s();

          case 15:
            if ((_step = _iterator.n()).done) {
              _context.next = 30;
              break;
            }

            idx = _step.value;
            countFilesToMove += 1;
            source = $("#" + idx.id + " a")[0].href.toString();
            target = DestinationUrl + $("#" + idx.id + " a")[0].href.split("/")[$("#" + idx.id + " a")[0].href.split("/").length - 1].toString();
            CopyFileName = source.split("/").pop();
            targetUrl = "/" + target.split(".com/")[1];
            ofile = web.getFileByServerRelativeUrl(targetUrl);
            context.load(ofile);
            FileFound = "";
            _context.next = 27;
            return new Promise(function (resolve, reject) {
              context.executeQueryAsync(function (x) {
                FileFound = "found";
                resolve();
              }, function (error) {
                console.log(error);
                resolve();
              });
            });

          case 27:
            if (FileFound == "found") {
              //Populate the File Exists list on the Copy Window
              document.getElementById("modalFileExists").style.display = "block";
              document.getElementById('modalFileExistList').innerHTML += CopyFileName + "</br>";
            } else {
              //Add the file to the copy util
              SP.MoveCopyUtil.copyFile(context, source, target, true);
              transferCount += 1;
            }

          case 28:
            _context.next = 15;
            break;

          case 30:
            _context.next = 35;
            break;

          case 32:
            _context.prev = 32;
            _context.t0 = _context["catch"](13);

            _iterator.e(_context.t0);

          case 35:
            _context.prev = 35;

            _iterator.f();

            return _context.finish(35);

          case 38:
            if (transferCount > 0) {
              context.executeQueryAsync(function () {
                document.getElementById('dvShowMessage').innerHTML = "<font color ='green'>Transfer complete</font>";
              }, function (sender, error) {
                var message = error.get_message();

                if (message.includes("file already exists")) {
                  document.getElementById("modalFileExists").style.display = "block";
                  document.getElementById('dvShowMessage').innerHTML = "";

                  var requestBody = sender.get_webRequest()._body;

                  var bodyXml = $.parseXML(requestBody);
                  $xml = $(bodyXml);
                  var sourceFileName = $xml.find("Parameter")[0].innerHTML.split("/").pop();
                  document.getElementById('modalFileExistList').innerHTML += sourceFileName + "</br>";
                } else {
                  document.getElementById("modalCopyError").style.display = "block";
                  document.getElementById('dvShowMessage').innerHTML = "";
                }
              });
            } else {
              document.getElementById('dvShowMessage').innerHTML = "";
            }

          case 39:
          case "end":
            return _context.stop();
        }
      }
    }, _callee, null, [[13, 32, 35, 38]]);
  }));
  return _copyFiles.apply(this, arguments);
}
</script>