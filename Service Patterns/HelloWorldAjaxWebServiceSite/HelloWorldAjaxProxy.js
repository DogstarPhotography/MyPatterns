
var helloWorldProxy;

// Initializes global and proxy default variables.
function pageLoad() {
	// Instantiate the service proxy.
	helloWorldProxy = new HelloWorldAjaxWebServiceSite.WebService.HelloWorld();

	// Set the default call back functions.
	helloWorldProxy.set_defaultSucceededCallback(SucceededCallback);
	helloWorldProxy.set_defaultFailedCallback(FailedCallback);
}


// Processes the button click and calls
// the service Greetings method.  
function OnClickGreetings() {
	var greetings = helloWorldProxy.Greetings();
}

// Callback function that
// processes the service return value.
function SucceededCallback(result) {
	var RsltElem = document.getElementById("Results");
	RsltElem.innerHTML = result;
}

// Callback function invoked when a call to 
// the  service methods fails.
function FailedCallback(error, userContext, methodName) {
	if (error !== null) {
		var RsltElem = document.getElementById("Results");

		RsltElem.innerHTML = "An error occurred: " +
            error.get_message();
	}
}

if (typeof (Sys) !== "undefined") Sys.Application.notifyScriptLoaded();