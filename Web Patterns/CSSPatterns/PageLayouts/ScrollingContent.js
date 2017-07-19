/*
Simple solution at last!
See http://stackoverflow.com/a/3730577/1374474
*/

// Scroll to element in div without moving the page
function ScrollToElementInDiv(id) {
	var item = document.getElementById(id);
	var docPos = GetDocPos(); // Get current position
	item.scrollIntoView(); // Scroll target into view
	window.scrollTo(0, docPos); // Scroll page back to original position
}

function GetDocPos() {
	return FilterResults(
        window.pageYOffset ? window.pageYOffset : 0,
        document.documentElement ? document.documentElement.scrollTop : 0,
        document.body ? document.body.scrollTop : 0
    );
}

function FilterResults(nWindow, nDocEl, nBody) {
	var nResult = nWindow ? nWindow : 0;
	if (nDocEl && (!nResult || (nResult > nDocEl)))
		nResult = nDocEl;
	return nBody && (!nResult || (nResult > nBody)) ? nBody : nResult;
}
