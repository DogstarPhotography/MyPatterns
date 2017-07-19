//
// Mini AJAX library
//

// Cross browser compatible XMLHttpRequest
function GetXMLHttpRequest() {
	var xmlhttp;
	if (window.XMLHttpRequest) {
		// code for IE7+, Firefox, Chrome, Opera, Safari
		xmlhttp = new XMLHttpRequest();
	}
	else {
		// code for IE6, IE5
		xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	}
	return xmlhttp;
}

// Synchronous requests
function RequestGet(url) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.open('GET', url, false);
	return xmlhttp.send();
}
function RequestPost(url, vars) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.open('POST', url, false);
	xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	return xmlhttp.send(vars);
}

// Asynchronous requests
function RequestGetAsync(url, postbackfunc) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			postbackfunc(xmlhttp.responseText);
		}
	}
	xmlhttp.open('GET', url, true);
	xmlhttp.send();
}
function RequestPostAsync(url, vars, postbackfunc) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			postbackfunc(xmlhttp.responseText);
		}
	}
	xmlhttp.open('POST', url, true);
	xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	xmlhttp.send(vars);
}

// Asynchronous requests for XML
function RequestGetAsyncXML(url, postbackfunc) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			postbackfunc(xmlhttp.responseXML);
		}
	}
	xmlhttp.open('GET', url, true);
	xmlhttp.send();
}
function RequestPostAsyncXML(url, vars, postbackfunc) {
	var xmlhttp = GetXMLHttpRequest();
	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			postbackfunc(xmlhttp.responseXML);
		}
	}
	xmlhttp.open('POST', url, true);
	xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	xmlhttp.send(vars);
}
