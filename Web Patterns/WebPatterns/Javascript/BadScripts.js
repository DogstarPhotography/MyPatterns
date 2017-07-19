// Doesn't work because getElementsByTagName returns only the element (<...>) not the node (<...>...</...>)
function ValidateQuestionnairePopoutMenu() {
	//Clear Navigation
	ClearPopout();
	//	<div id="PopoutMenu">
	//		<ul class="top-level">
	AppendPopout("<ul class=\"top-level\">");
	//Read document for navigation 'pages'
	var pages = document.getElementsByTagName('mbl:page');
	for (pgidx = 0; pgidx < pages.length; pgidx++) {
		var pgid = pages[pgidx].getAttribute('id');
		var pgtext = pages[pgidx].getAttribute('text');
		if (pgtext == null) { pgtext = pgid };
		if (pgtext.length == 0) { pgtext = pgid };
		//<li><a href="#">Home</a>
		//<ul class="sub-level">
		AppendPopout("<li><a href=\"#" + pgid + "\" >" + pgtext + "</a>");
		AppendPopout("<ul class=\"sub-level\">");
		//Read page for questions
		var questions = pages[pgidx].getElementsByTagName('mbl:question');
		//var questions = pages[pgidx].childNodes
		for (qidx = 0; qidx < questions.length; qidx++) {
			var qid = questions[qidx].getAttribute('id');
			var qtext = questions[qidx].getAttribute('text');
			if (qtext == null) { qtext = qid };
			if (qtext.length == 0) { qtext = qid };
			var actval = questions[qidx].getAttribute('value');
			var actvis = questions[qidx].getAttribute('visible');
			var actman = questions[qidx].getAttribute('mandatory');

			//<li><a href="#">Sub Menu Item 1</a></li>
			if (actvis == 1 && actman == 1) {
				if (actval.length > 0) {
					AppendPopout("<li><a href=\"#" + qid + "\" >" + qtext + "</a></li>")
				} else {
					AppendPopout("<li>X <a href=\"#" + qid + "\" >" + qtext + "</a></li>")
				}
			}
		};
		//</ul>
		//</li>
		AppendPopout("</ul></li>");
	};
	//</ul>
	AppendNavigation("</ul>");
};
