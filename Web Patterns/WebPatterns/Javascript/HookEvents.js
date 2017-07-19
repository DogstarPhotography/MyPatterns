
// This will add a function to a control event without removing any other functions already attached
function AddFunctionToControlEvent(TheControl, TheEventName, TheFunction) {
	// Check to see if the addEventListener function is available
	if (TheControl.addEventListener) 
		TheControl.addEventListener(TheEventName, TheFunction, false);
	// For browsers that do not support addEventListener use attachEvent
	else if (window.attachEvent) 
		TheControl.attachEvent(TheEventName, TheFunction);
}

// This will add a function to the window.onload event without messing up other functions
function AddOnload(myfunc) {
	if (window.addEventListener)
		window.addEventListener('load', myfunc, false);
	else if (window.attachEvent)
		window.attachEvent('onload', myfunc);
}

// This will hook all the control events without interfering with other events already attached
function HookEvents() {
	// Get all the inputs we want to add events to, generally this will be input fields
	var inputs = document.getinputsByTagName('input');
	for (i = 0; i < inputs.length; i++) {
		//Add a function call to the onblur event
		AddFunctionToControlEvent(inputs[i], 'onblur', HookedFunction);
	}
	// We can also monitor dropdowns
	var selects = document.getinputsByTagName('select');
	for (i = 0; i < selects.length; i++) {
		if (selects[i].addEventListener) {
			selects[i].addEventListener('onblur', HookedFunction, false);  
		}
		else {
			selects[i].attachEvent('onblur', HookedFunction);  
		}
	}
}

// This is the function that we will add to the events
function HookedFunction() {
	// Do something here
	alert('Hooked function called.');
}
