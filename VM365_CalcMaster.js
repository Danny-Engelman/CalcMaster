/*global _spPageContextInfo */
/*global document,alert,console,clearTimeout,setTimeout,window,XMLHttpRequest,$ */
/*jshint -W043*/

/**
 * Protect window.console method calls, e.g. console is not defined on IE
 * unless dev tools are open, and IE doesn't define console.debug
 *
 * Chrome 41.0.2272.118: debug,error,info,log,warn,dir,dirxml,table,trace,assert,count,markTimeline,profile,profileEnd,time,timeEnd,timeStamp,timeline,timelineEnd,group,groupCollapsed,groupEnd,clear
 * Firefox 37.0.1: log,info,warn,error,exception,debug,table,trace,dir,group,groupCollapsed,groupEnd,time,timeEnd,profile,profileEnd,assert,count
 * Internet Explorer 11: select,log,info,warn,error,debug,assert,time,timeEnd,timeStamp,group,groupCollapsed,groupEnd,trace,clear,dir,dirxml,count,countReset,cd
 * Safari 6.2.4: debug,error,log,info,warn,clear,dir,dirxml,table,trace,assert,count,profile,profileEnd,time,timeEnd,timeStamp,group,groupCollapsed,groupEnd
 * Opera 28.0.1750.48: debug,error,info,log,warn,dir,dirxml,table,trace,assert,count,markTimeline,profile,profileEnd,time,timeEnd,timeStamp,timeline,timelineEnd,group,groupCollapsed,groupEnd,clear
 */
(function () {
	// Union of Chrome, Firefox, IE, Opera, and Safari console methods
	var methods = ["assert", "assert", "cd", "clear", "count", "countReset",
    "debug", "dir", "dirxml", "dirxml", "dirxml", "error", "error", "exception",
    "group", "group", "groupCollapsed", "groupCollapsed", "groupEnd", "info",
    "info", "log", "log", "markTimeline", "profile", "profileEnd", "profileEnd",
    "select", "table", "table", "time", "time", "timeEnd", "timeEnd", "timeEnd",
    "timeEnd", "timeEnd", "timeStamp", "timeline", "timelineEnd", "trace",
    "trace", "trace", "trace", "trace", "warn"];
	var length = methods.length;
	var console = (window.console = window.console || {});
	var method;
	var noop = function () {};
	while (length--) {
		method = methods[length];
		// define undefined methods as noops to prevent errors
		if (!console[method])
			console[method] = noop;
	}
})();

if (typeof _spPageContextInfo !== undefined) {
	alert('Run the CalcMaster on a SharePoint page editing an existing Calculated Column Formula');
}
/* Global functions */
var VM365_CalcMaster = { //global obect for easier debugging
	//global settings
	comma: _spPageContextInfo.currentLanguage === 1033 ? ',' : ';', //SharePoint function separator
	layouts: _spPageContextInfo.layoutsUrl,
	saveID: 0,
	//initialize CalcMaster
	initialize: function () {
		//get references to DOM elements
		VM365_CalcMaster.columnList = document.getElementById('onetidIOCalcFields1');
		VM365_CalcMaster.textarea = document.getElementById('onetidIODefTextValue1');
		//adjust layout of textarea and column list
		VM365_CalcMaster.columnList.style.height = VM365_CalcMaster.textarea.style.height = '300px';
		VM365_CalcMaster.textarea.style.width = '500px';
		//Highlight the Number datatype for HTML output
		VM365_CalcMaster.setDOMelement('Number (outputs HTML/JavaScript)', 0, 'L_onetidTypeNumber1', 'lightGreen');
		//add CalcMaster UI elements for Toolbar and Save Formula Reporting, remember the MIT license
		var H = "<span style='font-size:1.3em;'>CalcMaster</span>";
		H += "powered by <a target=_new href=http://365coach.nl>365Coach.nl</a>";
		H += "<br>MIT sourcecode on <a target=_new href=https://github.com/Danny-Engelman/CalcMaster>GitHub</a>";
		H += "<br><br>See <a href='http://viewmaster365.com/365coach/#/Calculated_Column_Functions_List'";
		H += "target=_new>SharePoint Calculated Column Functions & Syntax List</a>";
		H += "<br>The";
		H += "<a href='http://viewmaster365.com/365coach/#/Drag_Drop_Columns_in_the_EditView_Page'>";
		H += "Drag-Drop View Editor</a> is a cool bookmarklet also";
		H += "<div style='min-height:55px;max-width:500px;'>";
		H += "<img src='/" + VM365_CalcMaster.layouts + "/images/WSC16.GIF'";
		H += "title='Wrap Javascript code in SharePoint Formula quoted string'";
		H += "onclick='VM365_CalcMaster.functionConvertToICC()'>";
		H += "<div id='calcmasterSave'></div><div id='calcmasterHint'></div>";
		H += "</div>";
		VM365_CalcMaster.setDOMelement(H, 0, 'onetidIODefText0');
		VM365_CalcMaster.addFunctionlist();
		VM365_CalcMaster.setDOMelement('Update your Formula and I will check if the Formula is correct', 2);
		//add event to the existing textarea
		VM365_CalcMaster.textarea.onkeyup = function (event) {
			if ([37, 38, 39, 40].indexOf(event.keyCode) === -1) { //ignore arrow keys
				clearTimeout(VM365_CalcMaster.saveID);
				VM365_CalcMaster.saveID = setTimeout(function () {
					VM365_CalcMaster.updateFormula(event);
				}, 300); //do not save while typing
			}
		};
	},
	//supporting functions
	setDOMelement: function (txt, color, id, bgcolor) {
		var element = document.getElementById(id || 'calcmasterSave');
		element.innerHTML = txt;
		element.style.color = ['black', 'red', 'green', 'darkorange'][color];
		if (bgcolor) element.style.backgroundColor = bgcolor;
	},
	setFormulaInTextarea: function (Formula) {
		if (Formula.indexOf('\n') === -1) { //if no newlines exist reformat to multiple lines
			Formula = Formula.replace(/&\"/gi, '\n  &\"'); //replace & before quote
			Formula = Formula.replace(/\"&/gi, '"\n  &'); //replace & after quote
			Formula = Formula.replace(/\";/gi, '\"\n;'); //replace & before quote
			Formula = Formula.replace(/IF/g, '\nIF'); //replace & before quote
		}
		VM365_CalcMaster.textarea.value = Formula;
	},
	sanitizeFormulaBeforeWritingToSP: function (Formula) {
		//the original Grand CalcMaster does a lot more, left the code in as a reminder to self
		//Formula = Formula.replace(/'/g, "&#39;");
		//Formula = Formula.replace(/'/g, "\\'");
		//Formula = Formula.replace(/'/g, '"&CHAR(39)&"');
		Formula = Formula.replace(/`/g, "'"); //replace backticks with single quote
		return (Formula);
	},
	//additional button functions in CalcMaster editor
	functionConvertToICC: function (Formula) {
		Formula = Formula || VM365_CalcMaster.textarea.value;
		if (Formula[0] !== '=') { //if Formula is NOT a SharePoint formula, it is javascript code
			Formula.replace(/'/g, '`'); //replace original single quotes with backticks (used in Grand CalcMaster)
			if (Formula.indexOf('onload') === -1 && Formula.indexOf('onclick') === -1) {
				//wrap in IMG onload tag and quoted SharePoint formula strings
				Formula = '="<img src=/_layouts/images/blank.gif onload=""{"&"' + Formula.replace(/\n/g, '"\n&"') + '"&"}"">"';
			}
			VM365_CalcMaster.setFormulaInTextarea(Formula);
			VM365_CalcMaster.updateFormula();
		} else {
			VM365_CalcMaster.setDOMelement('textarea does not contain plain JavaScript', 1);
		}
	},
	updateFormula: function (event) {
		//current selected Formula output datatype
		var selectedType = Array.prototype.filter.call(document.getElementsByName('ResultType'), function (r) {
			if (r.checked) return (r);
		});
		//document.getElementById('onetidSaveItem').style.display = 'none'; //hide OK button 
		var Formula = VM365_CalcMaster.textarea.value;
		VM365_CalcMaster.setFormulaInTextarea(Formula);
		VM365_CalcMaster.updateCalculatedColumn({
			Title: document.getElementById('idColName').value,
			Description: document.getElementById('idDesc').value,
			Formula: Formula,
			type: selectedType[0].value //Text or Number or DateTime or Currency or Boolean
		});
		//reporting of Formula analysis
		var calcmasterHint = [];
		if (Formula.split('"').length % 2 === 0) calcmasterHint.push("unmatched double quotes");
		if (Formula.split("'").length % 2 === 0) calcmasterHint.push("unmatched single quotes");
		if (Formula.split('(').length !== Formula.split(')').length) calcmasterHint.push("unmatched () brackets");
		VM365_CalcMaster.setDOMelement(calcmasterHint.join('<br>'), 3, 'calcmasterHint');
	},
	updateCalculatedColumn: function (column, formula, list, field) {
		VM365_CalcMaster.setDOMelement("Updating Formula " + column.Title + "... <span style='color:gray'>(this may take a minute)</span>", 2);
		column.list = "Lists(guid'" + _spPageContextInfo.pageListId.replace(/[{}]/g, '') + "')";
		column.field = "Fields/getbytitle('" + column.Title + "')";
		var xmlhttp = new XMLHttpRequest();
		xmlhttp.onreadystatechange = function () {
			console.info(xmlhttp.status, xmlhttp);
			VM365_CalcMaster.setDOMelement('Updating all List Items...');
			if (xmlhttp.readyState == XMLHttpRequest.DONE) {
				if (xmlhttp.status == 500) {
					console.info(xmlhttp.status, xmlhttp.statusText, JSON.parse(xmlhttp.responseText).error.message.value);
					VM365_CalcMaster.setDOMelement(JSON.parse(xmlhttp.responseText).error.message.value, 1);
				} else {
					VM365_CalcMaster.setDOMelement('Formula ' + column.Title + ' saved!');
				}
			}
		};
		console.info('saving', column.Title, column.Formula);
		xmlhttp.open("POST", _spPageContextInfo.webAbsoluteUrl + "/_api/Web/ " + column.list + "/" + column.field, true);
		xmlhttp.setRequestHeader("X-HTTP-Method", "MERGE");
		xmlhttp.setRequestHeader("X-RequestDigest", document.getElementById('__REQUESTDIGEST').value);
		xmlhttp.setRequestHeader("accept", "application/json;odata=verbose");
		xmlhttp.setRequestHeader("content-type", "application/json;odata=verbose");
		xmlhttp.send(
			JSON.stringify({
				'Title': column.Title,
				'Formula': VM365_CalcMaster.sanitizeFormulaBeforeWritingToSP(column.Formula),
				'__metadata': {
					'type': 'SP.FieldCalculated'
				},
				'Description': column.Description,
				'OutputType': ({
					Text: 2,
					Number: 9,
					Currency: 10,
					DateTime: 4,
					Boolean: 8
				})[column.type]
			})
		);
	},
	//leaving jQuery ajax call in here, I use it my CalcMaster Pro version
	JQ_updateCalculatedColumn: function (column) {
		VM365_CalcMaster.setDOMelement("Updating Formula " + column.Title + "... <span style='color:gray'>(this may take a minute)</span>", 2);
		column.list = "Lists(guid'" + _spPageContextInfo.pageListId.replace(/[{}]/g, '') + "')";
		column.field = "Fields/getbytitle('" + column.Title + "')";
		$.ajax({
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/ " + column.list + "/" + column.field,
			success: function (data) {
				VM365_CalcMaster.setDOMelement('Formula ' + column.Title + ' saved succesfully');
			},
			error: function (data) {
				VM365_CalcMaster.setDOMelement(data.responseJSON.error.message.value, 1);
			},
			type: "POST",
			headers: {
				"X-HTTP-Method": "MERGE",
				"X-RequestDigest": document.getElementById('__REQUESTDIGEST').value,
				"accept": "application/json;odata=verbose",
				"content-type": "application/json;odata=verbose"
			},
			data: JSON.stringify({
				'Title': column.Title,
				'Formula': VM365_CalcMaster.sanitizeFormulaBeforeWritingToSP(column.Formula),
				'__metadata': {
					'type': 'SP.FieldCalculated'
				},
				'Description': column.Description,
				'OutputType': ({
					Text: 2,
					Number: 9,
					Currency: 10,
					DateTime: 4,
					Boolean: 8
				})[column.type]
			})
		}); //ajax call
	},
	addFunctionlist: function () {
		var F = document.getElementById('onetidIODefText0').parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.cells[1];
		F.style.verticalAlign = 'top';
		F.style.width = 'inherit';
		F.innerHTML = "";
	}
};
VM365_CalcMaster.initialize();
